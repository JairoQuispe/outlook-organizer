const std = @import("std");
const ui = @import("../../../ui.zig");
const store_selector = @import("../../../store_selector.zig");
const types = @import("../types.zig");
const utils = @import("../utils.zig");
const scan_service = @import("../scan/scan_service.zig");
const routing_wizard = @import("../routing_wizard.zig");
const executor = @import("../executor.zig");
const import_config_factory = @import("import_config_factory.zig");
const summary_printer = @import("../summary/summary_printer.zig");
const shared_prompts_flow = @import("shared_prompts_flow.zig");
const batch_pst_config = @import("batch_pst_config.zig");
const shared_config_mod = @import("shared_config.zig");

const SharedConfig = shared_config_mod.SharedConfig;
const BatchPstConfig = batch_pst_config.BatchPstConfig;

const SelectedStore = struct {
    store_id: []const u8,
    store_name: []const u8,
    store_type: []const u8,
};

fn selectFirstAssignedStore(
    allocator: std.mem.Allocator,
    mappings: []const types.TargetStoreMapping,
) !SelectedStore {
    for (mappings) |m| {
        if (m.store_id.len == 0) continue;
        return .{
            .store_id = try allocator.dupe(u8, m.store_id),
            .store_name = try allocator.dupe(u8, m.store_name),
            .store_type = try allocator.dupe(u8, m.store_type),
        };
    }
    return error.NoStoreAssigned;
}

fn freeRoutingMappings(allocator: std.mem.Allocator, mappings: []const types.TargetStoreMapping) void {
    utils.freeTargetStoreMappings(allocator, mappings);
}

pub fn configureOnePst(
    allocator: std.mem.Allocator,
    pst_path: []const u8,
    shared: *const SharedConfig,
) !BatchPstConfig {
    const folder_selection = try scan_service.runScanAndSelectFolders(
        allocator,
        pst_path,
        shared.scan_mode,
        shared.scan_filter_year,
        shared.profile_name,
        shared.enable_routing,
    );

    var success = false;
    defer if (!success) {
        utils.cleanupTempFile(folder_selection.folder_plan_path);
        allocator.free(folder_selection.folder_plan_path);
        utils.cleanupTempFile(folder_selection.scan_export_path);
        allocator.free(folder_selection.scan_export_path);
    };

    var routing_mappings: ?[]const types.TargetStoreMapping = null;
    var selected_store: SelectedStore = .{ .store_id = "", .store_name = "", .store_type = "" };

    if (shared.enable_routing) {
        routing_mappings = try routing_wizard.configureMappings(
            allocator,
            shared.routing_criterion.?,
            folder_selection.scan_export_path,
            shared.profile_name,
        );

        selected_store = selectFirstAssignedStore(allocator, routing_mappings.?) catch {
            freeRoutingMappings(allocator, routing_mappings.?);
            return error.NoStoreAssigned;
        };
    } else {
        const chosen_store = try store_selector.selectTargetStore(allocator, shared.profile_name);
        selected_store = .{
            .store_id = chosen_store.store_id,
            .store_name = chosen_store.display_name,
            .store_type = chosen_store.store_type,
        };
    }

    success = true;
    return .{
        .pst_path = pst_path,
        .folder_plan_path = folder_selection.folder_plan_path,
        .scan_export_path = folder_selection.scan_export_path,
        .selected_count = folder_selection.selected_count,
        .total_count = folder_selection.total_count,
        .scan_mode = folder_selection.scan_mode,
        .target_store_id = selected_store.store_id,
        .target_store_name = selected_store.store_name,
        .target_store_type = selected_store.store_type,
        .routing_mappings = routing_mappings,
    };
}

pub fn runSingleImport(allocator: std.mem.Allocator, pst_path: []const u8) !void {
    var shared = shared_prompts_flow.promptSharedConfig(allocator, pst_path) catch |err| {
        if (err == error.Cancelled) {
            ui.cancelAbort();
            return;
        }
        ui.failAbort("Error en la configuracion inicial");
        return;
    };
    defer shared.deinit(allocator);

    var cfg = configureOnePst(allocator, pst_path, &shared) catch |err| {
        if (err == error.Cancelled) {
            ui.cancelAbort();
            return;
        }
        ui.failAbort("Error en la configuracion de la importacion");
        return;
    };
    defer cfg.deinit(allocator);

    const config = types.ImportConfig{
        .pst_path = pst_path,
        .target_store_id = cfg.target_store_id,
        .target_store_name = cfg.target_store_name,
        .target_store_type = cfg.target_store_type,
        .action = shared.action,
        .skip_duplicates = shared.skip_duplicates,
        .deep_duplicate_check = shared.deep_duplicate_check,
        .filter_year = shared.filter_year,
        .filter_months = shared.filter_months,
        .folder_plan_path = cfg.folder_plan_path,
        .adaptive_throttling = shared.adaptive_throttling,
        .profile_name = shared.profile_name,
        .routing_criterion = shared.routing_criterion,
        .routing_mappings = cfg.routing_mappings,
    };

    summary_printer.printSingleImportSummary(config, &shared, cfg.selected_count, cfg.total_count);

    std.debug.print("\n  \x1b[1;33mIniciar importacion? (S/n) [S]:\x1b[0m ", .{});
    const confirm = ui.readYesNo(true);

    if (!confirm) {
        std.debug.print("\n  \x1b[90mImportacion cancelada.\x1b[0m\n", .{});
        ui.waitForEnter();
        return;
    }

    try executor.executeImport(allocator, config);
}
