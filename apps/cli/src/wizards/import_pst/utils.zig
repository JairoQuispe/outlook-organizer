const std = @import("std");

pub fn extractNumber(json: []const u8, key: []const u8) ?i64 {
    var search_buf: [128]u8 = undefined;
    const search = std.fmt.bufPrint(&search_buf, "\"{s}\":", .{key}) catch return null;
    const key_pos = std.mem.indexOf(u8, json, search) orelse return null;
    var pos = key_pos + search.len;

    // Skip whitespace (just in case)
    while (pos < json.len and json[pos] == ' ') : (pos += 1) {}
    if (pos >= json.len) return null;

    // Read number (may have minus sign)
    var end = pos;
    if (end < json.len and json[end] == '-') end += 1;
    while (end < json.len and json[end] >= '0' and json[end] <= '9') : (end += 1) {}
    if (end == pos) return null;

    return std.fmt.parseInt(i64, json[pos..end], 10) catch null;
}

pub fn extractString(json: []const u8, key: []const u8) ?[]const u8 {
    var search_buf: [128]u8 = undefined;
    const search = std.fmt.bufPrint(&search_buf, "\"{s}\":\"", .{key}) catch return null;
    const start_pos = std.mem.indexOf(u8, json, search) orelse return null;
    const pos = start_pos + search.len;

    var end = pos;
    while (end < json.len) : (end += 1) {
        if (json[end] == '"' and (end == pos or json[end - 1] != '\\')) {
            return json[pos..end];
        }
    }
    return null;
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
