const std = @import("std");
const ui = @import("../../ui.zig");
const ps_runner = @import("../../ps_runner.zig");
const types = @import("types.zig");
const utils = @import("utils.zig");

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

pub fn onImportScriptLine(ctx: *anyopaque, line: []const u8) void {
    var state: *types.ImportProgressState = @ptrCast(@alignCast(ctx));
    if (line.len == 0) return;
    if (std.mem.indexOf(u8, line, "\"type\":\"progress\"") == null) return;

    if (utils.extractNumber(line, "copied")) |v| state.copied = v;
    if (utils.extractNumber(line, "moved")) |v| state.moved = v;
    if (utils.extractNumber(line, "skipped")) |v| state.skipped = v;
    if (utils.extractNumber(line, "failed")) |v| state.failed = v;
    if (utils.extractNumber(line, "sizeBytes")) |v| state.size_bytes = v;
    if (utils.extractNumber(line, "percent")) |v| {
        if (v <= 0) {
            state.percent = 0;
        } else if (v >= 100) {
            state.percent = 100;
        } else {
            state.percent = @intCast(v);
        }
    }

    const total_script_processed = state.copied + state.moved + state.skipped + state.failed;
    const elapsed_ms = std.time.milliTimestamp() - state.start_ms;

    if (state.has_rendered_progress) {
        // We now render 1 status line, 1 progress bar, and 6 metric lines = 8 lines total.
        // We go up 8 lines to overwrite the previous render cleanly.
        std.debug.print("\x1b[8F", .{});
    } else {
        state.has_rendered_progress = true;
    }

    const current_status = utils.extractString(line, "status") orelse "";
    const parsed = parseProcessedTotalFromStatus(current_status);

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

    // 1st line: Status / Folder
    std.debug.print("\r\x1b[2K  \x1b[90mCarpeta:\x1b[0m {s}\n", .{current_status});

    // 2nd line: Progress Bar (ui.printProgressBar prints \r\x1b[2K internally and appends no newline)
    ui.printProgressBar(percent_effective, "  Importando");

    const elapsed_sec = @divTrunc(elapsed_ms, 1000);
    const elapsed_parts = ui.secondsToHms(elapsed_sec);

    var size_buf: [32]u8 = undefined;
    var size_str: []const u8 = "0 Bytes";
    if (state.size_bytes > 0) {
        const size_gb = @as(f64, @floatFromInt(state.size_bytes)) / (1024.0 * 1024.0 * 1024.0);
        if (size_gb >= 0.1) {
            size_str = std.fmt.bufPrint(&size_buf, "{d:.2} GB", .{size_gb}) catch "Error";
        } else {
            const size_mb = @as(f64, @floatFromInt(state.size_bytes)) / (1024.0 * 1024.0);
            size_str = std.fmt.bufPrint(&size_buf, "{d:.2} MB", .{size_mb}) catch "Error";
        }
    }

    // 3rd to 8th line: Metric details (each has \x1b[2K to clear any leftover characters)
    std.debug.print("\n\x1b[2K  \x1b[90mProcesados:\x1b[0m     {d} ({s})", .{ processed_effective, size_str });
    std.debug.print("\n\x1b[2K  \x1b[90mCopiados:\x1b[0m        {d}", .{state.copied});
    std.debug.print("\n\x1b[2K  \x1b[90mMovidos:\x1b[0m         {d}", .{state.moved});
    std.debug.print("\n\x1b[2K  \x1b[90mOmitidos:\x1b[0m        {d}", .{state.skipped + state.failed});
    std.debug.print("\n\x1b[2K  \x1b[90mRestantes:\x1b[0m       {d}", .{remaining_effective});
    std.debug.print("\n\x1b[2K  \x1b[90mTranscurrido:\x1b[0m    {d:0>2}:{d:0>2}:{d:0>2}", .{
        elapsed_parts.hours,
        elapsed_parts.minutes,
        elapsed_parts.seconds,
    });

    var eta_str_buf: [32]u8 = undefined;
    var eta_str: []const u8 = "--:--:--";
    if (percent_effective > 0 and percent_effective < 100) {
        const total_est_ms = @divTrunc(elapsed_ms * 100, @as(i64, @intCast(percent_effective)));
        const eta_ms = @max(total_est_ms - elapsed_ms, 0);
        const eta_s = @divTrunc(eta_ms, 1000);
        const eta_parts = ui.secondsToHms(eta_s);
        eta_str = std.fmt.bufPrint(&eta_str_buf, "{d:0>2}:{d:0>2}:{d:0>2}", .{
            eta_parts.hours,
            eta_parts.minutes,
            eta_parts.seconds,
        }) catch "--:--:--";
    } else if (percent_effective >= 100) {
        eta_str = "00:00:00";
    }
    std.debug.print("  \x1b[90m| ETA:\x1b[0m {s}", .{eta_str});
    std.debug.print("\r", .{});
}

pub fn executeImport(allocator: std.mem.Allocator, config: types.ImportConfig) !void {
    ui.clearScreen();
    ui.printSectionTitle("Importando...");

    // Obtener tamaño del PST
    var pst_size_bytes: ?u64 = null;
    if (std.fs.openFileAbsolute(config.pst_path, .{})) |file| {
        defer file.close();
        if (file.stat()) |stat| {
            pst_size_bytes = stat.size;
        } else |_| {}
    } else |_| {}

    var size_str_buf: [32]u8 = undefined;
    var size_str: []const u8 = "Desconocido";
    if (pst_size_bytes) |bytes| {
        const size_gb = @as(f64, @floatFromInt(bytes)) / (1024.0 * 1024.0 * 1024.0);
        if (size_gb >= 0.1) {
            size_str = std.fmt.bufPrint(&size_str_buf, "{d:.2} GB", .{size_gb}) catch "Error";
        } else {
            const size_mb = @as(f64, @floatFromInt(bytes)) / (1024.0 * 1024.0);
            size_str = std.fmt.bufPrint(&size_str_buf, "{d:.2} MB", .{size_mb}) catch "Error";
        }
    }

    // Get start date and time
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

    // Mostrar el resumen general de los parámetros de importación
    std.debug.print("  \x1b[1;30m======================================================================\x1b[0m\n", .{});
    std.debug.print("  \x1b[1;37mPST de origen:\x1b[0m     {s} ({s})\n", .{ config.pst_path, size_str });
    const profile_display = if (config.profile_name) |p|
        if (p.len > 0) p else "Perfil predeterminado"
    else
        "Perfil predeterminado";
    std.debug.print("  \x1b[1;37mPerfil Outlook:\x1b[0m    {s}\n", .{profile_display});
    std.debug.print("  \x1b[1;37mBuzon de destino:\x1b[0m  {s}\n", .{config.target_store_name});

    var store_type_friendly: []const u8 = "Desconocido";
    if (std.mem.eql(u8, config.target_store_type, "ExchangeOnline")) {
        store_type_friendly = "Exchange Online";
    } else if (std.mem.eql(u8, config.target_store_type, "OST")) {
        store_type_friendly = "OST";
    } else if (std.mem.eql(u8, config.target_store_type, "PST")) {
        store_type_friendly = "PST";
    } else if (config.target_store_type.len > 0) {
        store_type_friendly = config.target_store_type;
    }
    const action_friendly = if (std.mem.eql(u8, config.action, "Move"))
        "Mover"
    else if (std.mem.eql(u8, config.action, "Copy"))
        "Copiar"
    else
        config.action;
    std.debug.print("  \x1b[1;37mTipo de buzon:\x1b[0m     {s}\n", .{store_type_friendly});
    std.debug.print("  \x1b[1;37mAccion:\x1b[0m          {s}\n", .{action_friendly});
    std.debug.print("  \x1b[1;37mInicio proceso:\x1b[0m    {s}\n", .{date_time_str});

    if (config.filter_year) |y| {
        std.debug.print("  \x1b[1;37mFiltro de anios:\x1b[0m  {s}\n", .{y});
    } else {
        std.debug.print("  \x1b[1;37mFiltro de anios:\x1b[0m  Todos los anios\n", .{});
    }
    std.debug.print("  \x1b[1;30m======================================================================\x1b[0m\n\n", .{});

    // Write the import script to temp
    const script_path = try ps_runner.writeEmbeddedScript(allocator, .import_pst);
    defer ps_runner.cleanupScript(allocator, script_path);

    // Build arguments
    var args = std.ArrayListUnmanaged([]const u8){};
    defer args.deinit(allocator);

    try args.appendSlice(allocator, &.{
        "-PstPath",
        config.pst_path,
        "-TargetStoreId",
        config.target_store_id,
        "-Action",
        config.action,
        "-Json",
        "-Headless",
    });

    if (config.skip_duplicates) {
        try args.append(allocator, "-SkipDuplicates");
    }
    if (config.deep_duplicate_check) {
        try args.append(allocator, "-DeepDuplicateCheck");
    }
    if (config.filter_year) |y| {
        try args.append(allocator, "-FilterOnlyYear");
        try args.append(allocator, y);
    }
    if (config.filter_months) |m| {
        try args.append(allocator, "-FilterOnlyMonths");
        try args.append(allocator, m);
    }
    try args.append(allocator, "-FolderPlanPath");
    try args.append(allocator, config.folder_plan_path);
    if (config.adaptive_throttling) {
        try args.append(allocator, "-AdaptiveThrottling");
    }
    if (config.profile_name) |p| {
        if (p.len > 0) {
            try args.append(allocator, "-ProfileName");
            try args.append(allocator, p);
        }
    }

    std.debug.print("  \x1b[90mEjecutando script de importacion...\x1b[0m\n\n", .{});
    const start_time = std.time.milliTimestamp();
    var progress_state = types.ImportProgressState{
        .start_ms = start_time,
        .copied = 0,
        .moved = 0,
        .skipped = 0,
        .failed = 0,
        .size_bytes = 0,
        .percent = 0,
        .has_rendered_progress = false,
    };

    const script_run = ps_runner.runScriptDetailedStreaming(allocator, script_path, args.items, onImportScriptLine, &progress_state) catch {
        ui.printError("Error ejecutando el script de importacion");
        ui.waitForEnter();
        return;
    };
    defer allocator.free(script_run.command_line);
    defer allocator.free(script_run.output);

    std.debug.print("  \x1b[90mComando:\x1b[0m {s}\n\n", .{script_run.command_line});
    std.debug.print("\n", .{});

    const output = script_run.output;

    const elapsed_total = std.time.milliTimestamp() - start_time;

    // Parse output: find progress and restoreResult lines
    var last_copied: i64 = 0;
    var last_moved: i64 = 0;
    var last_skipped: i64 = 0;
    var last_failed: i64 = 0;
    var result_line: ?[]const u8 = null;
    var error_message: ?[]const u8 = null;

    var line_iter = std.mem.splitSequence(u8, output, "\n");
    while (line_iter.next()) |raw_line| {
        const line = std.mem.trim(u8, raw_line, &[_]u8{ ' ', '\t', '\r', '\n' });
        if (line.len == 0) continue;
        if (std.mem.indexOf(u8, line, "\"type\"") == null) continue;

        if (std.mem.indexOf(u8, line, "\"progress\"") != null) {
            if (utils.extractNumber(line, "copied")) |c| last_copied = c;
            if (utils.extractNumber(line, "moved")) |m| last_moved = m;
            if (utils.extractNumber(line, "skipped")) |s| last_skipped = s;
            if (utils.extractNumber(line, "failed")) |f| last_failed = f;
        } else if (std.mem.indexOf(u8, line, "\"error\"") != null) {
            error_message = utils.extractString(line, "message") orelse line;
        } else if (std.mem.indexOf(u8, line, "\"restoreResult\"") != null) {
            result_line = line;
        }
    }

    if (error_message) |msg| {
        ui.printError(msg);
        ui.waitForEnter();
        return;
    }

    if (result_line == null and output.len == 0) {
        ui.printError("El script de importacion no devolvio salida. Revisa que Outlook este abierto y el PST sea accesible.");
        ui.waitForEnter();
        return;
    }

    if (result_line == null and last_copied == 0 and last_moved == 0 and last_skipped == 0 and last_failed == 0) {
        ui.printError("La importacion no devolvio resultado final (restoreResult).");
        if (output.len > 0) {
            std.debug.print("\n  \x1b[91mSalida/Error del script:\x1b[0m\n", .{});
            std.debug.print("  {s}\n", .{output});
        }
        ui.waitForEnter();
        return;
    }

    // Show final results
    std.debug.print("\n", .{});
    ui.printSectionTitle("Resultado de importacion");

    const json = result_line orelse "";
    if (json.len > 0) {
        const copied = utils.extractNumber(json, "copied") orelse last_copied;
        const moved = utils.extractNumber(json, "moved") orelse last_moved;
        const skipped = utils.extractNumber(json, "skipped") orelse last_skipped;
        const failed = utils.extractNumber(json, "failed") orelse last_failed;
        const throttle_events = utils.extractNumber(json, "throttleEvents") orelse 0;

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
            std.debug.print("\n  \x1b[1;31mHubo {d} fallos. Revisa los logs para mas detalle.\x1b[0m\n", .{failed});
        } else {
            ui.printSuccess("Importacion completada exitosamente.");
        }
    } else {
        std.debug.print("  \x1b[1;32mCopiados:\x1b[0m    {d}\n", .{last_copied});
        std.debug.print("  \x1b[1;32mMovidos:\x1b[0m     {d}\n", .{last_moved});
        std.debug.print("  \x1b[1;33mOmitidos:\x1b[0m    {d}\n", .{last_skipped});
        std.debug.print("  \x1b[1;31mFallidos:\x1b[0m    {d}\n", .{last_failed});
        std.debug.print("  \x1b[90mTiempo:\x1b[0m      {d}s\n", .{@divTrunc(elapsed_total, 1000)});
    }

    ui.waitForEnter();
}
