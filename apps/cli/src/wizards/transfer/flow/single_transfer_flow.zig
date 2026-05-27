const std = @import("std");
const ui = @import("../../../ui.zig");
const store_selector = @import("../../../store_selector.zig");
const file_browser = @import("../../../file_browser.zig");
const types = @import("../types.zig");
const source_selector = @import("../source_selector.zig");
const dest_selector = @import("../dest_selector.zig");
const shared_prompts = @import("shared_prompts_flow.zig");
const transfer_runner = @import("../execution/transfer_runner.zig");
const scan_service = @import("../../import_pst/scan/scan_service.zig");
const folder_selector = @import("../../import_pst/folder_selector.zig");
const scan_output_parser = @import("../../import_pst/scan/scan_output_parser.zig");
const import_types = @import("../../import_pst/types.zig");

pub fn runSingleTransfer(allocator: std.mem.Allocator) !void {
    // Step 1: Select source
    const source = try source_selector.selectSource(allocator, null);
    defer {
        allocator.free(source.store_id);
        allocator.free(source.store_name);
        allocator.free(source.store_type);
        allocator.free(source.pst_path);
    }

    // Step 2: Scan source and select folders
    const source_is_pst = source.pst_path.len > 0;
    const source_identifier = if (source_is_pst)
        source.pst_path
    else
        source.store_id;

    var folder_selection = scanFolders(allocator, source_identifier, source_is_pst) catch {
        ui.failAbort("Error escaneando origen o seleccionando carpetas");
        return;
    };
    defer {
        folder_selector.cleanupFolderSelection(allocator, &folder_selection);
    }

    // Step 3: Select destination
    const dest = try dest_selector.selectDestination(allocator, null);
    defer {
        allocator.free(dest.store_id);
        allocator.free(dest.store_name);
        allocator.free(dest.store_type);
        allocator.free(dest.pst_path);
    }

    // Step 4: Shared config (profile, action, duplicates, filters, throttling)
    var shared = shared_prompts.promptSharedConfig(allocator) catch {
        ui.failAbort("Error en la configuracion");
        return;
    };
    defer shared.deinit(allocator);

    // Step 5: Build config and print summary
    const config = types.TransferConfig{
        .source_info = types.SourceInfo{
            .store_id = try allocator.dupe(u8, source.store_id),
            .store_name = try allocator.dupe(u8, source.store_name),
            .store_type = try allocator.dupe(u8, source.store_type),
            .pst_path = try allocator.dupe(u8, source.pst_path),
        },
        .dest_info = types.DestInfo{
            .store_id = try allocator.dupe(u8, dest.store_id),
            .store_name = try allocator.dupe(u8, dest.store_name),
            .store_type = try allocator.dupe(u8, dest.store_type),
            .pst_path = try allocator.dupe(u8, dest.pst_path),
        },
        .action = try allocator.dupe(u8, shared.action),
        .skip_duplicates = shared.skip_duplicates,
        .deep_duplicate_check = shared.deep_duplicate_check,
        .filter_year = if (shared.filter_year) |y| try allocator.dupe(u8, y) else null,
        .filter_months = if (shared.filter_months) |m| try allocator.dupe(u8, m) else null,
        .folder_plan_path = try allocator.dupe(u8, folder_selection.folder_plan_path),
        .adaptive_throttling = shared.adaptive_throttling,
        .profile_name = if (shared.profile_name) |p| try allocator.dupe(u8, p) else null,
        .routing_criterion = null,
        .routing_mappings = null,
    };
    defer {
        allocator.free(config.source_info.store_id);
        allocator.free(config.source_info.store_name);
        allocator.free(config.source_info.store_type);
        allocator.free(config.source_info.pst_path);
        allocator.free(config.dest_info.store_id);
        allocator.free(config.dest_info.store_name);
        allocator.free(config.dest_info.store_type);
        allocator.free(config.dest_info.pst_path);
        allocator.free(config.action);
        if (config.filter_year) |y| allocator.free(y);
        if (config.filter_months) |m| allocator.free(m);
        allocator.free(config.folder_plan_path);
        if (config.profile_name) |p| allocator.free(p);
    }

    // Step 6: Print summary and confirm
    printTransferSummary(config, folder_selection.selected_count, folder_selection.total_count);

    std.debug.print("\n  \x1b[1;33mIniciar transferencia? (S/n) [S]:\x1b[0m ", .{});
    const confirm = ui.readYesNo(true);
    if (!confirm) {
        std.debug.print("\n  \x1b[90mTransferencia cancelada.\x1b[0m\n", .{});
        ui.waitForEnter();
        return;
    }

    // Step 7: Execute
    try transfer_runner.executeTransfer(allocator, config);
}

fn scanFolders(allocator: std.mem.Allocator, identifier: []const u8, is_pst: bool) !import_types.FolderSelectionResult {
    _ = is_pst;
    // Delegate to the import scan service which supports both PstPath and StoreId
    return try scan_service.runScanWithSource(
        allocator,
        identifier,
        .quick,
        null,
        null,
        false,
    );
}

pub fn printTransferSummary(config: types.TransferConfig, folder_count: usize, folder_total: usize) void {
    ui.clearScreen();
    ui.printSectionTitle("Resumen de transferencia");

    const src_is_pst = config.source_info.pst_path.len > 0;
    if (src_is_pst) {
        std.debug.print("  \x1b[1;37mOrigen PST:\x1b[0m     {s}\n", .{config.source_info.store_name});
    } else {
        std.debug.print("  \x1b[1;37mOrigen:\x1b[0m         {s} ({s})\n", .{ config.source_info.store_name, config.source_info.store_type });
    }

    const dst_is_pst = config.dest_info.pst_path.len > 0;
    if (dst_is_pst) {
        std.debug.print("  \x1b[1;37mDestino PST:\x1b[0m    {s}\n", .{config.dest_info.pst_path});
    } else {
        std.debug.print("  \x1b[1;37mDestino:\x1b[0m        {s} ({s})\n", .{ config.dest_info.store_name, config.dest_info.store_type });
    }

    std.debug.print("  \x1b[1;37mAccion:\x1b[0m        {s}\n", .{if (std.mem.eql(u8, config.action, "Move")) "Mover" else "Copiar"});
    std.debug.print("  \x1b[1;37mCarpetas:\x1b[0m      {d}/{d} seleccionadas\n", .{ folder_count, folder_total });
    std.debug.print("  \x1b[1;37mDuplicados:\x1b[0m    {s}\n", .{if (config.skip_duplicates) "Saltar" else "No saltar"});
    if (config.deep_duplicate_check) {
        std.debug.print("  \x1b[1;37mRevision:\x1b[0m      Profunda\n", .{});
    }
    if (config.filter_year) |y| {
        std.debug.print("  \x1b[1;37mAnio:\x1b[0m          {s}\n", .{y});
    }
    if (config.filter_months) |m| {
        std.debug.print("  \x1b[1;37mMeses:\x1b[0m         {s}\n", .{m});
    }
    std.debug.print("  \x1b[1;37mThrottling:\x1b[0m    {s}\n", .{if (config.adaptive_throttling) "Adaptativo" else "Fijo"});
}
