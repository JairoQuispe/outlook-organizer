const std = @import("std");
const ui = @import("../../../ui.zig");
const ps_runner = @import("../../../ps_runner.zig");
const types = @import("../types.zig");
const import_utils = @import("../../import_pst/utils.zig");

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

fn ensureProcessSafeArgs(allocator: std.mem.Allocator, args: []const []const u8) ![]const []const u8 {
    const out = try allocator.alloc([]const u8, args.len);
    for (args, 0..) |arg, i| {
        if (std.unicode.utf8ValidateSlice(arg)) {
            out[i] = arg;
        } else {
            out[i] = try sanitizeProcessArgAscii(allocator, arg);
        }
    }
    return out;
}

fn sanitizeProcessArgAscii(allocator: std.mem.Allocator, raw: []const u8) ![]u8 {
    var out = std.ArrayListUnmanaged(u8){};
    errdefer out.deinit(allocator);

    for (raw) |ch| {
        if ((ch >= 0x20 and ch <= 0x7E) or ch == '\t') {
            try out.append(allocator, ch);
        }
    }

    return try out.toOwnedSlice(allocator);
}

pub fn buildTransferArgs(allocator: std.mem.Allocator, config: types.TransferConfig) ![]const []const u8 {
    var builder = ArgsBuilder.init(allocator);
    errdefer builder.deinit();

    try builder.addSlice(&.{ "-Json", "-Headless" });

    if (config.source_info.pst_path.len > 0) {
        try builder.addOption("-SourcePstPath", config.source_info.pst_path);
    } else if (config.source_info.store_id.len > 0) {
        const source_store_id = try sanitizeStoreId(allocator, config.source_info.store_id);
        if (source_store_id.len == 0) return error.InvalidStoreId;
        try builder.addOption("-SourceStoreId", source_store_id);
    }

    if (config.dest_info.pst_path.len > 0) {
        try builder.addOption("-DestPstPath", config.dest_info.pst_path);
    } else if (config.dest_info.store_id.len > 0) {
        const dest_store_id = try sanitizeStoreId(allocator, config.dest_info.store_id);
        if (dest_store_id.len == 0) return error.InvalidStoreId;
        try builder.addOption("-DestStoreId", dest_store_id);
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
        const sanitized_store_id = try sanitizeStoreId(allocator, m.store_id);
        defer allocator.free(sanitized_store_id);
        if (sanitized_store_id.len == 0) continue;
        if (!first) try json.appendSlice(allocator, ",");
        first = false;

        try json.appendSlice(allocator, "{\"storeId\":\"");
        try appendJsonEscaped(allocator, &json, sanitized_store_id);
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

    if (first) return error.NoValidMappings;

    return try json.toOwnedSlice(allocator);
}

fn sanitizeStoreId(allocator: std.mem.Allocator, raw: []const u8) ![]u8 {
    var out = std.ArrayListUnmanaged(u8){};
    errdefer out.deinit(allocator);

    for (raw) |ch| {
        if ((ch >= '0' and ch <= '9') or
            (ch >= 'a' and ch <= 'f') or
            (ch >= 'A' and ch <= 'F'))
        {
            try out.append(allocator, std.ascii.toUpper(ch));
        }
    }

    return try out.toOwnedSlice(allocator);
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

    var source_size_bytes: ?u64 = null;
    if (config.source_info.pst_path.len > 0) {
        if (std.fs.openFileAbsolute(config.source_info.pst_path, .{})) |file| {
            defer file.close();
            if (file.stat()) |stat| {
                source_size_bytes = stat.size;
            } else |_| {}
        } else |_| {}
    }

    var size_buf: [32]u8 = undefined;
    const source_size_str: []const u8 = if (source_size_bytes) |bytes|
        import_utils.formatBytesShortFromU64(&size_buf, bytes)
    else
        "Desconocido";

    const SYSTEMTIME = extern struct {
        wYear: u16,
        wMonth: u16,
        wDayOfWeek: u16,
        wDay: u16,
        wHour: u16,
        wMinute: u16,
        wSecond: u16,
        wMilliseconds: u16,
    };
    const kernel32 = struct {
        extern "kernel32" fn GetLocalTime(lpSystemTime: *SYSTEMTIME) void;
    };
    var start_sys_time: SYSTEMTIME = undefined;
    kernel32.GetLocalTime(&start_sys_time);

    var date_time_buf: [64]u8 = undefined;
    const date_time_str = std.fmt.bufPrint(&date_time_buf, "{0d:0>2}/{1d:0>2}/{2d:0>4} {3d:0>2}:{4d:0>2}:{5d:0>2}", .{
        start_sys_time.wDay,
        start_sys_time.wMonth,
        start_sys_time.wYear,
        start_sys_time.wHour,
        start_sys_time.wMinute,
        start_sys_time.wSecond,
    }) catch "Desconocida";

    printTransferExecutionSummary(config, source_size_str, date_time_str);

    const script_path = try ps_runner.writeEmbeddedScript(allocator, .transfer);
    defer ps_runner.cleanupScript(allocator, script_path);

    var args_arena = std.heap.ArenaAllocator.init(allocator);
    defer args_arena.deinit();
    const args = try buildTransferArgs(args_arena.allocator(), config);
    const safe_args = try ensureProcessSafeArgs(args_arena.allocator(), args);

    std.debug.print("  \x1b[90mEjecutando script de transferencia...\x1b[0m\n\n", .{});

    const start_time = std.time.milliTimestamp();
    var progress_state = TransferProgressState{
        .start_ms = start_time,
        .copied = 0,
        .moved = 0,
        .skipped = 0,
        .failed = 0,
        .size_bytes = 0,
        .percent = 0,
        .has_rendered_progress = false,
    };

    const script_run = ps_runner.runScriptDetailedStreaming(allocator, script_path, safe_args, onTransferScriptLine, &progress_state) catch |err| {
        if (progress_state.has_rendered_progress) {
            std.debug.print("\n", .{});
        }
        ui.printError("Error ejecutando el script de transferencia");
        std.debug.print("  \x1b[90mDetalle:\x1b[0m {s}\n", .{@errorName(err)});
        const command_preview = ps_runner.buildCommandPreview(allocator, "powershell.exe", script_path, safe_args) catch null;
        if (command_preview) |cmd| {
            defer allocator.free(cmd);
            std.debug.print("  \x1b[90mComando:\x1b[0m {s}\n", .{cmd});
        }
        ui.waitForEnter();
        return;
    };
    defer allocator.free(script_run.command_line);
    defer allocator.free(script_run.output);

    if (progress_state.has_rendered_progress) {
        std.debug.print("\n", .{});
    }

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

fn printTransferExecutionSummary(config: types.TransferConfig, source_size_str: []const u8, date_time_str: []const u8) void {
    std.debug.print("  \x1b[1;30m======================================================================\x1b[0m\n", .{});

    if (config.source_info.pst_path.len > 0) {
        std.debug.print("  \x1b[1;37mOrigen PST:\x1b[0m      {s} ({s})\n", .{ config.source_info.pst_path, source_size_str });
    } else {
        std.debug.print("  \x1b[1;37mOrigen buzon:\x1b[0m    {s}\n", .{config.source_info.store_name});
    }

    std.debug.print("  \x1b[1;37mPerfil Outlook:\x1b[0m   {s}\n", .{import_utils.profileDisplayName(config.profile_name)});

    if (config.routing_criterion) |criterion| {
        std.debug.print("  \x1b[1;37mEnrutamiento:\x1b[0m     Multibuzon (agrupado por {s})\n", .{if (criterion == .by_year) "Anos" else "Meses"});
        std.debug.print("  \x1b[1;37mMapeos asignados:\x1b[0m {d}\n", .{countAssignedTransferMappings(config.routing_mappings)});
    } else if (config.dest_info.pst_path.len > 0) {
        std.debug.print("  \x1b[1;37mDestino PST:\x1b[0m     {s}\n", .{config.dest_info.pst_path});
    } else {
        std.debug.print("  \x1b[1;37mBuzon destino:\x1b[0m   {s}\n", .{config.dest_info.store_name});
        std.debug.print("  \x1b[1;37mTipo de buzon:\x1b[0m    {s}\n", .{import_utils.storeTypeDisplayName(config.dest_info.store_type)});
    }

    std.debug.print("  \x1b[1;37mAccion:\x1b[0m          {s}\n", .{if (std.mem.eql(u8, config.action, "Move")) "Mover" else "Copiar"});
    std.debug.print("  \x1b[1;37mDuplicados:\x1b[0m      {s}\n", .{if (config.skip_duplicates) "Saltar" else "No saltar"});
    std.debug.print("  \x1b[1;37mRevision dup:\x1b[0m    {s}\n", .{if (config.deep_duplicate_check) "Profunda" else "Simple"});
    std.debug.print("  \x1b[1;37mThrottling:\x1b[0m     {s}\n", .{if (config.adaptive_throttling) "Adaptativo" else "Fijo"});
    std.debug.print("  \x1b[1;37mCarpetas plan:\x1b[0m   {s}\n", .{config.folder_plan_path});

    if (config.filter_year) |y| {
        std.debug.print("  \x1b[1;37mFiltro de anios:\x1b[0m {s}\n", .{y});
    } else {
        std.debug.print("  \x1b[1;37mFiltro de anios:\x1b[0m Todos los anios\n", .{});
    }
    if (config.filter_months) |m| {
        std.debug.print("  \x1b[1;37mFiltro de meses:\x1b[0m {s}\n", .{m});
    }

    std.debug.print("  \x1b[1;37mInicio proceso:\x1b[0m   {s}\n", .{date_time_str});
    std.debug.print("  \x1b[1;30m======================================================================\x1b[0m\n\n", .{});
}

fn countAssignedTransferMappings(mappings: ?[]const types.TargetStoreMapping) usize {
    const items = mappings orelse return 0;
    var count: usize = 0;
    for (items) |m| {
        if (m.store_id.len > 0) count += 1;
    }
    return count;
}

pub const TransferProgressState = types.TransferProgressState;

const ParsedProgress = struct {
    processed: i64,
    total: i64,
};

fn parseProcessedTotalFromStatus(status: []const u8) ?ParsedProgress {
    const close_idx = std.mem.lastIndexOfScalar(u8, status, ')') orelse return null;
    const open_idx = std.mem.lastIndexOfScalar(u8, status[0..close_idx], '(') orelse return null;
    if (open_idx >= close_idx) return null;

    const inner = std.mem.trim(u8, status[open_idx + 1 .. close_idx], " ");
    const slash_idx = std.mem.indexOfScalar(u8, inner, '/') orelse return null;
    const left = std.mem.trim(u8, inner[0..slash_idx], " ");
    const right = std.mem.trim(u8, inner[slash_idx + 1 ..], " ");
    if (left.len == 0 or right.len == 0) return null;

    const processed = std.fmt.parseInt(i64, left, 10) catch return null;
    const total = std.fmt.parseInt(i64, right, 10) catch return null;
    if (processed < 0 or total <= 0) return null;

    return .{ .processed = processed, .total = total };
}

pub fn onTransferScriptLine(ctx: *anyopaque, line: []const u8) void {
    const state: *TransferProgressState = @ptrCast(@alignCast(ctx));
    parseProgressLine(state, line);
}

fn parseProgressLine(state: *TransferProgressState, line: []const u8) void {
    const kind = extractJsonString(line, "type") orelse return;
    if (!std.mem.eql(u8, kind, "progress")) return;

    const status = extractJsonString(line, "status") orelse "";
    state.copied = extractJsonNumber(line, "copied") orelse state.copied;
    state.moved = extractJsonNumber(line, "moved") orelse state.moved;
    state.skipped = extractJsonNumber(line, "skipped") orelse state.skipped;
    state.failed = extractJsonNumber(line, "failed") orelse state.failed;
    state.size_bytes = extractJsonNumber(line, "sizeBytes") orelse state.size_bytes;

    const raw_percent = extractJsonNumber(line, "percent") orelse @as(i64, @intCast(state.percent));
    if (raw_percent <= 0) {
        state.percent = 0;
    } else if (raw_percent >= 100) {
        state.percent = 100;
    } else {
        state.percent = @intCast(raw_percent);
    }

    const total_script_processed = state.copied + state.moved + state.skipped + state.failed;
    const elapsed_ms = std.time.milliTimestamp() - state.start_ms;

    if (state.has_rendered_progress) {
        std.debug.print("\x1b[7F", .{});
    } else {
        state.has_rendered_progress = true;
    }

    const parsed = parseProcessedTotalFromStatus(status);
    const processed_effective = if (parsed) |p|
        @max(p.processed, total_script_processed)
    else
        total_script_processed;

    const remaining_effective = if (parsed) |p|
        @max(p.total - processed_effective, 0)
    else blk: {
        const percent_fallback: u32 = if (state.percent > 0)
            if (state.percent >= 100) 100 else state.percent
        else
            @as(u32, if (processed_effective > 0) 1 else 0);

        var estimated_total = processed_effective;
        if (percent_fallback > 0) {
            var est = @divTrunc(processed_effective * 100, @as(i64, @intCast(percent_fallback)));
            if (est < processed_effective) est = processed_effective;
            estimated_total = est;
        }

        break :blk @max(estimated_total - processed_effective, 0);
    };

    const percent_effective: u32 = if (parsed) |p| blk: {
        const capped_processed = @min(processed_effective, p.total);
        var pct = @as(u32, @intCast(@divTrunc(capped_processed * 100, p.total)));
        if (pct > 99 and capped_processed < p.total) pct = 99;
        break :blk pct;
    } else if (state.percent > 0)
        if (state.percent >= 100) 100 else state.percent
    else
        @as(u32, if (processed_effective > 0) 1 else 0);

    const columns = ui.terminalWidthColumns();
    const status_prefix_len = "  Carpeta: ".len;
    const status_max = if (columns > status_prefix_len) columns - status_prefix_len else 0;
    var status_buf: [512]u8 = undefined;
    const safe_status = ui.truncateWithEllipsis(status, &status_buf, status_max);

    std.debug.print("\r\x1b[2K  \x1b[90mCarpeta:\x1b[0m {s}\n", .{safe_status});
    ui.printProgressBar(percent_effective, "Transfiriendo");

    var size_buf: [32]u8 = undefined;
    const size_str = import_utils.formatBytesShort(&size_buf, state.size_bytes);

    std.debug.print("\n\x1b[2K  \x1b[90mProc:\x1b[0m {d} ({s})", .{ processed_effective, size_str });
    std.debug.print("\n\x1b[2K  \x1b[90mCop:\x1b[0m  {d}", .{state.copied});
    std.debug.print("\n\x1b[2K  \x1b[90mMov:\x1b[0m  {d}", .{state.moved});
    std.debug.print("\n\x1b[2K  \x1b[90mOmi:\x1b[0m  {d}", .{state.skipped + state.failed});
    std.debug.print("\n\x1b[2K  \x1b[90mRes:\x1b[0m  {d}", .{remaining_effective});

    var elapsed_hms_buf: [32]u8 = undefined;
    const elapsed_hms = import_utils.formatHms(&elapsed_hms_buf, @divTrunc(elapsed_ms, 1000));
    std.debug.print("\n\x1b[2K  \x1b[90mT:\x1b[0m {s}", .{elapsed_hms});

    var eta_str_buf: [32]u8 = undefined;
    var eta_str: []const u8 = "--:--:--";
    if (percent_effective > 0 and percent_effective < 100) {
        const total_est_ms = @divTrunc(elapsed_ms * 100, @as(i64, @intCast(percent_effective)));
        const eta_ms = @max(total_est_ms - elapsed_ms, 0);
        eta_str = import_utils.formatHms(&eta_str_buf, @divTrunc(eta_ms, 1000));
    } else if (percent_effective >= 100) {
        eta_str = "00:00:00";
    }

    std.debug.print("  \x1b[90mETA:\x1b[0m {s}", .{eta_str});
    std.debug.print("\r", .{});
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
        if (json[pos] == '\\') {
            pos += 2;
            continue;
        }
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
