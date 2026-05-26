const std = @import("std");
const ui = @import("../../../ui.zig");
const ps_runner = @import("../../../ps_runner.zig");
const types = @import("../types.zig");
const utils = @import("../utils.zig");
const args_builder_mod = @import("../args_builder.zig");
const progress_renderer = @import("progress_renderer.zig");
const result_parser = @import("result_parser.zig");
const result_summarizer = @import("../summary/summary_printer.zig");
const routing_payload_builder = @import("routing_payload_builder.zig");

pub fn executeImport(allocator: std.mem.Allocator, config: types.ImportConfig) !void {
    ui.clearScreen();
    ui.printSectionTitle("Importando...");

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
        size_str = utils.formatBytesShortFromU64(&size_str_buf, bytes);
    }

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

    result_summarizer.printImportSummary(config, size_str, date_time_str);

    const script_path = try ps_runner.writeEmbeddedScript(allocator, .import_pst);
    defer ps_runner.cleanupScript(allocator, script_path);

    var args_builder = args_builder_mod.ArgsBuilder.init(allocator);
    defer args_builder.deinit();

    try args_builder.addSlice(&.{
        "-PstPath",
        config.pst_path,
        "-TargetStoreId",
        config.target_store_id,
        "-Action",
        config.action,
        "-Json",
        "-Headless",
    });

    if (config.routing_criterion) |criterion| {
        try args_builder.addOption("-RoutingCriterion", if (criterion == .by_year) "by_year" else "by_month");

        if (config.routing_mappings) |mappings| {
            const json = routing_payload_builder.buildRoutingMappingsJson(allocator, mappings) catch |err| {
                if (err == error.NoValidMappings) {
                    ui.failAbort("Enrutamiento activo sin mapeos validos (store_id vacio).");
                    return;
                }
                return err;
            };
            defer allocator.free(json);
            try args_builder.addOption("-RoutingMappingsJson", json);
        } else {
            ui.failAbort("Enrutamiento activo sin RoutingMappings configurados.");
            return;
        }
    }

    try args_builder.addBoolFlag(config.skip_duplicates, "-SkipDuplicates");
    try args_builder.addBoolFlag(config.deep_duplicate_check, "-DeepDuplicateCheck");
    if (config.filter_year) |y| {
        try args_builder.addOption("-FilterOnlyYear", y);
    }
    if (config.filter_months) |m| {
        try args_builder.addOption("-FilterOnlyMonths", m);
    }
    try args_builder.addOption("-FolderPlanPath", config.folder_plan_path);
    try args_builder.addBoolFlag(config.adaptive_throttling, "-AdaptiveThrottling");
    try args_builder.addOptionIfNonEmpty("-ProfileName", config.profile_name);

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

    const script_run = ps_runner.runScriptDetailedStreaming(allocator, script_path, args_builder.items(), progress_renderer.onImportScriptLine, &progress_state) catch {
        ui.failAbort("Error ejecutando el script de importacion");
        return;
    };
    defer allocator.free(script_run.command_line);
    defer allocator.free(script_run.output);

    std.debug.print("  \x1b[90mComando final:\x1b[0m {s}\n", .{script_run.command_line});
    std.debug.print("  \x1b[90mExit code:\x1b[0m {d}\n\n", .{script_run.exit_code});
    std.debug.print("\n", .{});

    const output = script_run.output;

    const elapsed_total = std.time.milliTimestamp() - start_time;

    if (script_run.exit_code != 0) {
        ui.printError("El script de importacion finalizo con error.");
        if (output.len > 0) {
            std.debug.print("\n  \x1b[91mSalida del script:\x1b[0m\n", .{});
            std.debug.print("  {s}\n", .{output});
        } else {
            std.debug.print("  \x1b[90mSin salida capturada en stdout.\x1b[0m\n", .{});
        }
        ui.waitForEnter();
        return;
    }

    const parsed_output = result_parser.parseImportOutput(output);

    if (parsed_output.error_message) |msg| {
        ui.failAbort(msg);
        return;
    }

    if (parsed_output.result_line == null and output.len == 0) {
        ui.failAbort("El script de importacion no devolvio salida. Revisa que Outlook este abierto y el PST sea accesible.");
        return;
    }

    if (parsed_output.result_line == null and parsed_output.last_copied == 0 and parsed_output.last_moved == 0 and parsed_output.last_skipped == 0 and parsed_output.last_failed == 0) {
        ui.printError("La importacion no devolvio resultado final (restoreResult).");
        if (output.len > 0) {
            std.debug.print("\n  \x1b[91mSalida/Error del script:\x1b[0m\n", .{});
            std.debug.print("  {s}\n", .{output});
        }
        ui.waitForEnter();
        return;
    }

    result_summarizer.printImportResult(&parsed_output, elapsed_total);
    ui.waitForEnter();
}
