const std = @import("std");
const types = @import("types.zig");

fn isWhitespace(c: u8) bool {
    return c == ' ' or c == '\t' or c == '\r' or c == '\n';
}

fn findJsonValueStart(json: []const u8, key: []const u8) ?usize {
    var search_buf: [128]u8 = undefined;
    const search = std.fmt.bufPrint(&search_buf, "\"{s}\"", .{key}) catch return null;
    const key_pos = std.mem.indexOf(u8, json, search) orelse return null;
    var pos = key_pos + search.len;

    while (pos < json.len and isWhitespace(json[pos])) : (pos += 1) {}
    if (pos >= json.len or json[pos] != ':') return null;
    pos += 1;

    while (pos < json.len and isWhitespace(json[pos])) : (pos += 1) {}
    if (pos >= json.len) return null;
    return pos;
}

pub fn extractNumber(json: []const u8, key: []const u8) ?i64 {
    const pos = findJsonValueStart(json, key) orelse return null;
    var end = pos;
    if (end < json.len and json[end] == '-') end += 1;
    while (end < json.len and json[end] >= '0' and json[end] <= '9') : (end += 1) {}
    if (end == pos) return null;

    return std.fmt.parseInt(i64, json[pos..end], 10) catch null;
}

pub fn extractString(json: []const u8, key: []const u8) ?[]const u8 {
    const pos = findJsonValueStart(json, key) orelse return null;
    if (json[pos] != '"') return null;
    var end = pos + 1;
    while (end < json.len) : (end += 1) {
        if (json[end] == '"' and (end == pos + 1 or json[end - 1] != '\\')) {
            return json[pos + 1 .. end];
        }
    }
    return null;
}

pub fn findEnclosingObject(buffer: []const u8, key_pos: usize) ?[]const u8 {
    const obj_start = std.mem.lastIndexOfScalar(u8, buffer[0 .. key_pos + 1], '{') orelse return null;
    const obj_end = std.mem.indexOfScalarPos(u8, buffer, key_pos, '}') orelse return null;
    if (obj_end <= obj_start) return null;
    return buffer[obj_start .. obj_end + 1];
}

pub fn extractYearBreakdownDisplay(allocator: std.mem.Allocator, json: []const u8) ![]u8 {
    const token = "\"yearBreakdown\":[";
    const start = std.mem.indexOf(u8, json, token) orelse return allocator.dupe(u8, "");

    var pos = start + token.len;
    var out = std.ArrayListUnmanaged(u8){};
    errdefer out.deinit(allocator);

    var wrote_any = false;
    while (pos < json.len) {
        if (json[pos] == ']') break;

        if (json[pos] == '{') {
            const obj_end_rel = std.mem.indexOfScalarPos(u8, json, pos, '}') orelse break;
            const row = json[pos .. obj_end_rel + 1];

            const year = extractNumber(row, "year");
            const count = extractNumber(row, "count");
            if (year != null and count != null) {
                if (wrote_any) try out.appendSlice(allocator, ", ");

                const part = try std.fmt.allocPrint(allocator, "{d}:{d}", .{ year.?, count.? });
                defer allocator.free(part);
                try out.appendSlice(allocator, part);
                wrote_any = true;
            }

            pos = obj_end_rel + 1;
            continue;
        }

        pos += 1;
    }

    if (!wrote_any) {
        out.deinit(allocator);
        return allocator.dupe(u8, "");
    }

    return try out.toOwnedSlice(allocator);
}

pub fn makeTempFilePath(allocator: std.mem.Allocator, prefix: []const u8, ext: []const u8) ![]u8 {
    const temp_dir = std.process.getEnvVarOwned(allocator, "TEMP") catch |err| switch (err) {
        error.EnvironmentVariableNotFound => try allocator.dupe(u8, "."),
        else => return err,
    };
    defer allocator.free(temp_dir);

    const pid: u32 = std.os.windows.GetCurrentProcessId();
    const ts: u64 = @intCast(@max(std.time.milliTimestamp(), 0));
    const file_name = try std.fmt.allocPrint(allocator, "{s}-{d}-{d}.{s}", .{ prefix, pid, ts, ext });
    defer allocator.free(file_name);

    return try std.fs.path.join(allocator, &.{ temp_dir, file_name });
}

pub fn cleanupTempFile(path: []const u8) void {
    if (std.fs.path.isAbsolute(path)) {
        std.fs.deleteFileAbsolute(path) catch {};
    } else {
        std.fs.cwd().deleteFile(path) catch {};
    }
}

pub fn storeTypeDisplayName(store_type: []const u8) []const u8 {
    if (std.mem.eql(u8, store_type, "ExchangeOnline")) return "Exchange Online";
    if (std.mem.eql(u8, store_type, "OST")) return "OST";
    if (std.mem.eql(u8, store_type, "PST")) return "PST";
    return if (store_type.len > 0) store_type else "Desconocido";
}

pub fn profileDisplayName(profile_name: ?[]const u8) []const u8 {
    return if (profile_name) |p|
        if (p.len > 0) p else "Perfil predeterminado"
    else
        "Perfil predeterminado";
}

pub fn countAssignedMappings(mappings: ?[]const types.TargetStoreMapping) usize {
    const items = mappings orelse return 0;
    var count: usize = 0;
    for (items) |m| {
        if (m.store_id.len > 0) count += 1;
    }
    return count;
}

pub fn routingCriterionDisplay(criterion: types.RoutingCriterion) []const u8 {
    return if (criterion == .by_year)
        "Múltiples buzones agrupados por Años"
    else
        "Múltiples buzones agrupados por Meses";
}

pub fn appendJsonEscaped(allocator: std.mem.Allocator, list: *std.ArrayListUnmanaged(u8), value: []const u8) !void {
    for (value) |ch| {
        switch (ch) {
            '"' => try list.appendSlice(allocator, "\\\""),
            '\\' => try list.appendSlice(allocator, "\\\\"),
            '\n' => try list.appendSlice(allocator, "\\n"),
            '\r' => try list.appendSlice(allocator, "\\r"),
            '\t' => try list.appendSlice(allocator, "\\t"),
            else => try list.append(allocator, ch),
        }
    }
}

pub fn freeTargetStoreMappings(allocator: std.mem.Allocator, mappings: []const types.TargetStoreMapping) void {
    for (mappings) |m| {
        m.deinit(allocator);
    }
    allocator.free(mappings);
}

pub fn formatBytesShort(buf: []u8, bytes: i64) []const u8 {
    if (bytes <= 0) return "0 Bytes";

    const gb = @as(f64, @floatFromInt(bytes)) / (1024.0 * 1024.0 * 1024.0);
    if (gb >= 0.1) {
        return std.fmt.bufPrint(buf, "{d:.2} GB", .{gb}) catch "Error";
    }

    const mb = @as(f64, @floatFromInt(bytes)) / (1024.0 * 1024.0);
    return std.fmt.bufPrint(buf, "{d:.2} MB", .{mb}) catch "Error";
}

pub fn formatBytesShortFromU64(buf: []u8, bytes: u64) []const u8 {
    const i = std.math.cast(i64, bytes) orelse return "Error";
    return formatBytesShort(buf, i);
}

pub fn formatHms(buf: []u8, total_seconds: i64) []const u8 {
    const safe = @max(total_seconds, 0);
    const hours = @divTrunc(safe, 3600);
    const rem = @rem(safe, 3600);
    const minutes = @divTrunc(rem, 60);
    const seconds = @rem(rem, 60);
    return std.fmt.bufPrint(buf, "{d:0>2}:{d:0>2}:{d:0>2}", .{ hours, minutes, seconds }) catch "--:--:--";
}
