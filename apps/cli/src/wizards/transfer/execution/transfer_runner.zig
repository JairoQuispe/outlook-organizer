const std = @import("std");
const ui = @import("../../../ui.zig");
const ps_runner = @import("../../../ps_runner.zig");
const types = @import("../types.zig");

pub const ArgsBuilder = struct {
    allocator: std.mem.Allocator,
    list: std.ArrayListUnmanaged([]const u8),

    pub fn init(allocator: std.mem.Allocator) ArgsBuilder {
        return .{ .allocator = allocator, .list = .{} };
    }

    pub fn deinit(self: *ArgsBuilder) void {
        self.list.deinit(self.allocator);
    }

    pub fn items(self: *const ArgsBuilder) []const []const u8 {
        return self.list.items;
    }

    pub fn addSlice(self: *ArgsBuilder, values: []const []const u8) !void {
        try self.list.appendSlice(self.allocator, values);
    }

    pub fn addFlag(self: *ArgsBuilder, flag: []const u8) !void {
        try self.list.append(self.allocator, flag);
    }

    pub fn addOption(self: *ArgsBuilder, flag: []const u8, value: []const u8) !void {
        try self.list.append(self.allocator, flag);
        try self.list.append(self.allocator, value);
    }

    pub fn addBoolFlag(self: *ArgsBuilder, enabled: bool, flag: []const u8) !void {
        if (!enabled) return;
        try self.addFlag(flag);
    }

    pub fn addOptionIfNonEmpty(self: *ArgsBuilder, flag: []const u8, value: ?[]const u8) !void {
        const v = value orelse return;
        if (v.len == 0) return;
        try self.addOption(flag, v);
    }
};

pub fn buildTransferArgs(allocator: std.mem.Allocator, config: types.TransferConfig) ![]const []const u8 {
    var builder = ArgsBuilder.init(allocator);
    errdefer builder.deinit();

    try builder.addSlice(&.{"-Json", "-Headless"});

    if (config.source_info.pst_path.len > 0) {
        try builder.addOption("-SourcePstPath", config.source_info.pst_path);
    } else if (config.source_info.store_id.len > 0) {
        try builder.addOption("-SourceStoreId", config.source_info.store_id);
    }

    if (config.dest_info.pst_path.len > 0) {
        try builder.addOption("-DestPstPath", config.dest_info.pst_path);
    } else if (config.dest_info.store_id.len > 0) {
        try builder.addOption("-DestStoreId", config.dest_info.store_id);
    }

    try builder.addOption("-Action", config.action);
    try builder.addBoolFlag(config.skip_duplicates, "-SkipDuplicates");
    try builder.addBoolFlag(config.deep_duplicate_check, "-DeepDuplicateCheck");
    try builder.addBoolFlag(config.adaptive_throttling, "-AdaptiveThrottling");

    if (config.filter_year) |y| {
        try builder.addOption("-FilterOnlyYear", y);
    }
    if (config.filter_months) |m| {
        try builder.addOption("-FilterOnlyMonths", m);
    }
    if (config.folder_plan_path.len > 0) {
        try builder.addOption("-FolderPlanPath", config.folder_plan_path);
    }
    try builder.addOptionIfNonEmpty("-ProfileName", config.profile_name);

    if (config.routing_criterion) |criterion| {
        try builder.addOption("-RoutingCriterion", if (criterion == .by_year) "by_year" else "by_month");

        if (config.routing_mappings) |mappings| {
            const json = try buildRoutingMappingsJson(allocator, mappings);
            defer allocator.free(json);
            try builder.addOption("-RoutingMappingsJson", json);
        }
    }

    const result = try allocator.alloc([]const u8, builder.items().len);
    @memcpy(result, builder.items());
    return result;
}

fn buildRoutingMappingsJson(allocator: std.mem.Allocator, mappings: []const types.TargetStoreMapping) ![]u8 {
    var json = std.ArrayListUnmanaged(u8){};
    errdefer json.deinit(allocator);

    try json.appendSlice(allocator, "[");
    var first = true;
    for (mappings) |m| {
        if (m.store_id.len == 0) continue;
        if (!first) try json.appendSlice(allocator, ",");
        first = false;

        try json.appendSlice(allocator, "{\"storeId\":\"");
        try appendJsonEscaped(allocator, &json, m.store_id);
        try json.appendSlice(allocator, "\",\"year\":");
        const y = try std.fmt.allocPrint(allocator, "{d}", .{m.year});
        defer allocator.free(y);
        try json.appendSlice(allocator, y);

        if (m.month) |mo| {
            try json.appendSlice(allocator, ",\"month\":");
            const mo_str = try std.fmt.allocPrint(allocator, "{d}", .{mo});
            defer allocator.free(mo_str);
            try json.appendSlice(allocator, mo_str);
        }

        try json.appendSlice(allocator, "}");
    }
    try json.appendSlice(allocator, "]");

    return try json.toOwnedSlice(allocator);
}

fn appendJsonEscaped(allocator: std.mem.Allocator, list: *std.ArrayListUnmanaged(u8), value: []const u8) !void {
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

pub const ParsedTransferOutput = struct {
    last_copied: i64 = 0,
    last_moved: i64 = 0,
    last_skipped: i64 = 0,
    last_failed: i64 = 0,
    result_line: ?[]const u8 = null,
    error_message: ?[]const u8 = null,
};

pub fn parseTransferOutput(output: []const u8) ParsedTransferOutput {
    var parsed = ParsedTransferOutput{};
    var line_iter = std.mem.splitSequence(u8, output, "\n");
    while (line_iter.next()) |raw_line| {
        const line = std.mem.trim(u8, raw_line, &[_]u8{ ' ', '\t', '\r', '\n' });
        if (line.len == 0) continue;

        const kind = extractJsonString(line, "type") orelse continue;

        if (std.mem.eql(u8, kind, "progress")) {
            if (extractJsonNumber(line, "copied")) |v| parsed.last_copied = v;
            if (extractJsonNumber(line, "moved")) |v| parsed.last_moved = v;
            if (extractJsonNumber(line, "skipped")) |v| parsed.last_skipped = v;
            if (extractJsonNumber(line, "failed")) |v| parsed.last_failed = v;
        } else if (std.mem.eql(u8, kind, "error")) {
            parsed.error_message = extractJsonString(line, "message");
        } else if (std.mem.eql(u8, kind, "restoreResult")) {
            parsed.result_line = line;
        }
    }
    return parsed;
}

pub fn executeTransfer(allocator: std.mem.Allocator, config: types.TransferConfig) !void {
    ui.clearScreen();
    ui.printSectionTitle("Transfiriendo...");

    const script_path = try ps_runner.writeEmbeddedScript(allocator, .transfer);
    defer ps_runner.cleanupScript(allocator, script_path);

    const args = try buildTransferArgs(allocator, config);
    defer allocator.free(args);

    std.debug.print("  \x1b[90mEjecutando script de transferencia...\x1b[0m\n\n", .{});

    const start_time = std.time.milliTimestamp();

    const script_run = ps_runner.runScriptDetailedStreaming(allocator, script_path, args, onTransferScriptLine, undefined) catch {
        ui.failAbort("Error ejecutando el script de transferencia");
        return;
    };
    defer allocator.free(script_run.command_line);
    defer allocator.free(script_run.output);

    std.debug.print("  \x1b[90mExit code:\x1b[0m {d}\n\n", .{script_run.exit_code});

    const output = script_run.output;
    const elapsed_total = std.time.milliTimestamp() - start_time;

    if (script_run.exit_code != 0) {
        ui.printError("El script finalizo con error.");
        if (output.len > 0) {
            std.debug.print("\n  \x1b[91mSalida:\x1b[0m\n", .{});
            std.debug.print("  {s}\n", .{output});
        }
        ui.waitForEnter();
        return;
    }

    const parsed = parseTransferOutput(output);

    if (parsed.error_message) |msg| {
        ui.failAbort(msg);
        return;
    }

    printTransferResult(&parsed, elapsed_total);
    ui.waitForEnter();
}

pub const TransferProgressState = types.TransferProgressState;

pub fn onTransferScriptLine(ctx: *anyopaque, line: []const u8) void {
    _ = ctx;
    _ = parseProgressLine(line);
}

fn parseProgressLine(line: []const u8) void {
    const kind = extractJsonString(line, "type") orelse return;
    if (!std.mem.eql(u8, kind, "progress")) return;

    const status = extractJsonString(line, "status") orelse return;
    const percent = extractJsonNumber(line, "percent") orelse 0;
    const copied = extractJsonNumber(line, "copied") orelse 0;
    const moved = extractJsonNumber(line, "moved") orelse 0;
    const skipped = extractJsonNumber(line, "skipped") orelse 0;
    const failed = extractJsonNumber(line, "failed") orelse 0;

    std.debug.print("\r\x1b[2K  \x1b[90m{s}\x1b[0m {d}% | C:{d} M:{d} S:{d} F:{d}", .{ status, percent, copied, moved, skipped, failed });
}

pub fn printTransferResult(parsed: *const ParsedTransferOutput, elapsed_total: i64) void {
    std.debug.print("\n", .{});
    ui.printSectionTitle("Resultado de transferencia");

    const json = parsed.result_line orelse "";
    if (json.len > 0) {
        const copied = extractJsonNumber(json, "copied") orelse parsed.last_copied;
        const moved = extractJsonNumber(json, "moved") orelse parsed.last_moved;
        const skipped = extractJsonNumber(json, "skipped") orelse parsed.last_skipped;
        const failed = extractJsonNumber(json, "failed") orelse parsed.last_failed;
        const throttle_events = extractJsonNumber(json, "throttleEvents") orelse 0;

        const elapsed_sec = @divTrunc(elapsed_total, 1000);
        const elapsed_min = @divTrunc(elapsed_sec, 60);
        const elapsed_s = @rem(elapsed_sec, 60);

        std.debug.print("  \x1b[1;32mCopiados:\x1b[0m         {d}\n", .{copied});
        std.debug.print("  \x1b[1;32mMovidos:\x1b[0m          {d}\n", .{moved});
        std.debug.print("  \x1b[1;33mOmitidos:\x1b[0m         {d}\n", .{skipped});
        std.debug.print("  \x1b[1;31mFallidos:\x1b[0m         {d}\n", .{failed});
        std.debug.print("  \x1b[90mTiempo:\x1b[0m           {d}:{d:0>2}\n", .{ elapsed_min, elapsed_s });
        std.debug.print("  \x1b[90mThrottle eventos:\x1b[0m {d}\n", .{throttle_events});

        const total = copied + moved + skipped + failed;
        if (total > 0) {
            std.debug.print("\n  \x1b[1;36mTotal procesados: {d}\x1b[0m\n", .{total});
        }

        if (failed > 0) {
            std.debug.print("\n  \x1b[1;31mHubo {d} fallos.\x1b[0m\n", .{failed});
        } else {
            ui.printSuccess("Transferencia completada exitosamente.");
        }
    } else {
        std.debug.print("  \x1b[1;32mCopiados:\x1b[0m    {d}\n", .{parsed.last_copied});
        std.debug.print("  \x1b[1;32mMovidos:\x1b[0m     {d}\n", .{parsed.last_moved});
        std.debug.print("  \x1b[1;33mOmitidos:\x1b[0m    {d}\n", .{parsed.last_skipped});
        std.debug.print("  \x1b[1;31mFallidos:\x1b[0m    {d}\n", .{parsed.last_failed});
        std.debug.print("  \x1b[90mTiempo:\x1b[0m      {d}s\n", .{@divTrunc(elapsed_total, 1000)});
    }
}

fn extractJsonString(json: []const u8, key: []const u8) ?[]const u8 {
    var search_buf: [128]u8 = undefined;
    const search = std.fmt.bufPrint(&search_buf, "\"{s}\"", .{key}) catch return null;
    const key_pos = std.mem.indexOf(u8, json, search) orelse return null;
    var pos = key_pos + search.len;

    while (pos < json.len and (json[pos] == ':' or json[pos] == ' ' or json[pos] == '\t' or json[pos] == '\n' or json[pos] == '\r')) : (pos += 1) {}

    if (pos >= json.len or json[pos] != '"') return null;
    pos += 1;

    const value_start = pos;
    while (pos < json.len) {
        if (json[pos] == '\\') { pos += 2; continue; }
        if (json[pos] == '"') return json[value_start..pos];
        pos += 1;
    }
    return null;
}

fn extractJsonNumber(json: []const u8, key: []const u8) ?i64 {
    var search_buf: [128]u8 = undefined;
    const search = std.fmt.bufPrint(&search_buf, "\"{s}\"", .{key}) catch return null;
    const key_pos = std.mem.indexOf(u8, json, search) orelse return null;
    var pos = key_pos + search.len;

    while (pos < json.len and (json[pos] == ':' or json[pos] == ' ' or json[pos] == '\t' or json[pos] == '\n' or json[pos] == '\r')) : (pos += 1) {}

    if (pos >= json.len) return null;
    if (json[pos] == '"') return null;

    var end = pos;
    if (end < json.len and json[end] == '-') end += 1;
    while (end < json.len and json[end] >= '0' and json[end] <= '9') : (end += 1) {}
    if (end == pos) return null;

    return std.fmt.parseInt(i64, json[pos..end], 10) catch null;
}
