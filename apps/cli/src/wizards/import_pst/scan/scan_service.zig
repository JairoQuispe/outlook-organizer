const std = @import("std");
const ui = @import("../../../ui.zig");
const ps_runner = @import("../../../ps_runner.zig");
const types = @import("../types.zig");
const utils = @import("../utils.zig");
const folder_selector = @import("../folder_selector.zig");
const args_builder_mod = @import("../args_builder.zig");
const scan_output_parser = @import("scan_output_parser.zig");

pub fn runScanAndSelectFolders(
    allocator: std.mem.Allocator,
    pst_path: []const u8,
    scan_mode: types.ScanMode,
    scan_filter_year: ?[]const u8,
    profile_name: ?[]const u8,
    enable_routing: bool,
) !types.FolderSelectionResult {
    return runScanWithSource(allocator, pst_path, scan_mode, scan_filter_year, profile_name, enable_routing);
}

pub fn runScanWithSource(
    allocator: std.mem.Allocator,
    source_id: []const u8,
    scan_mode: types.ScanMode,
    scan_filter_year: ?[]const u8,
    profile_name: ?[]const u8,
    enable_routing: bool,
) !types.FolderSelectionResult {
    // Auto-detect: if source looks like an absolute path, it's a PST file
    const is_pst = std.fs.path.isAbsolute(source_id);

    ui.clearScreen();
    if (is_pst) {
        ui.printSectionTitle("Escanear PST");
    } else {
        ui.printSectionTitle("Escanear origen");
    }
    std.debug.print("  \x1b[90mEjecutando escaneo {s}...\x1b[0m\n", .{if (scan_mode == .deep) "profundo" else "rapido"});

    const scan_script_path = try ps_runner.writeEmbeddedScript(allocator, .scan_pst);
    defer ps_runner.cleanupScript(allocator, scan_script_path);

    const exported_scan_path = try utils.makeTempFilePath(allocator, "oo-scan-export", "json");
    defer allocator.free(exported_scan_path);

    var args_builder = args_builder_mod.ArgsBuilder.init(allocator);
    defer args_builder.deinit();

    if (is_pst) {
        try args_builder.addSlice(&.{ "-PstPath", source_id });
    } else {
        try args_builder.addSlice(&.{ "-StoreId", source_id });
    }

    try args_builder.addSlice(&.{
        "-Json",
        "-Headless",
        "-ExportFolders",
        "-ExportFoldersPath",
        exported_scan_path,
    });
    if (scan_mode == .deep or enable_routing) {
        try args_builder.addFlag("-IncludeSize");
    }
    if (enable_routing) {
        try args_builder.addFlag("-ExportStatistics");
    }
    if (scan_filter_year) |y| {
        try args_builder.addOption("-FilterOnlyYear", y);
    }
    try args_builder.addOptionIfNonEmpty("-ProfileName", profile_name);

    const output = ps_runner.runScript(allocator, scan_script_path, args_builder.items()) catch return error.ScanFailed;
    defer allocator.free(output);

    var folders = scan_output_parser.parseScannedFolders(allocator, output) catch return error.ScanFailed;
    defer {
        for (folders.items) |folder| {
            allocator.free(folder.path);
            allocator.free(folder.year_breakdown_display);
        }
        folders.deinit(allocator);
    }

    if (folders.items.len == 0) {
        ui.printError("No se encontraron carpetas con items para importar");
        return error.NoFoldersFound;
    }

    const selected_flags = try folder_selector.promptFolderSelection(allocator, folders.items, scan_mode);
    defer allocator.free(selected_flags);

    const selected_count = folder_selector.countSelectedFlags(selected_flags);
    if (selected_count == 0) {
        return error.Cancelled;
    }

    const plan_path = try folder_selector.writeFolderPlanFromFlags(allocator, folders.items, selected_flags);
    const scan_export_copy = try allocator.dupe(u8, exported_scan_path);

    return .{
        .folder_plan_path = plan_path,
        .scan_export_path = scan_export_copy,
        .selected_count = selected_count,
        .total_count = folders.items.len,
        .scan_mode = scan_mode,
    };
}
