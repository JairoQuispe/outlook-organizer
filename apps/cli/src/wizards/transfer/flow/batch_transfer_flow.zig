const std = @import("std");
const ui = @import("../../../ui.zig");
const types = @import("../types.zig");
const source_selector = @import("../source_selector.zig");
const dest_selector = @import("../dest_selector.zig");
const shared_prompts = @import("shared_prompts_flow.zig");
const transfer_runner = @import("../execution/transfer_runner.zig");
const scan_service = @import("../../import_pst/scan/scan_service.zig");
const folder_selector = @import("../../import_pst/folder_selector.zig");
const import_types = @import("../../import_pst/types.zig");

const BatchSourceConfig = struct {
    source: types.SourceInfo,
    folder_selection: import_types.FolderSelectionResult,

    fn deinit(self: *BatchSourceConfig, allocator: std.mem.Allocator) void {
        allocator.free(self.source.store_id);
        allocator.free(self.source.store_name);
        allocator.free(self.source.store_type);
        allocator.free(self.source.pst_path);
        folder_selector.cleanupFolderSelection(allocator, &self.folder_selection);
    }
};

pub fn runMultiSourceBatch(allocator: std.mem.Allocator) !void {
    ui.clearScreen();
    ui.printSectionTitle("Transferencia por Lote");
    std.debug.print("  \x1b[90mSelecciona los origenes uno por uno, luego el destino comun.\x1b[0m\n", .{});
    std.debug.print("  \x1b[90mPresiona Q para finalizar la seleccion de origenes.\x1b[0m\n\n", .{});
    ui.waitForEnter();

    // Step 1: Collect sources
    var sources = std.ArrayListUnmanaged(BatchSourceConfig){};
    defer {
        for (sources.items) |*s| s.deinit(allocator);
        sources.deinit(allocator);
    }

    while (true) {
        ui.clearScreen();
        std.debug.print("  \x1b[1;33mOrigen #{d}\x1b[0m\n\n", .{sources.items.len + 1});
        std.debug.print("  \x1b[90mQ = finalizar seleccion de origenes\x1b[0m\n\n", .{});

        const src = source_selector.selectSource(allocator, null) catch |err| {
            if (err == error.Cancelled) {
                if (sources.items.len == 0) {
                    std.debug.print("\n  \x1b[90mOperacion cancelada.\x1b[0m\n", .{});
                    ui.waitForEnter();
                    return;
                }
                break;
            }
            ui.failAbort("Error seleccionando origen");
            return;
        };

        const is_pst = src.pst_path.len > 0;
        const identifier = if (is_pst) src.pst_path else src.store_id;

        const folder_sel = scanFolders(allocator, identifier, is_pst) catch |err| {
            ui.printError("Error escaneando origen, se omite.");
            ui.waitForEnter();
            allocator.free(src.store_id);
            allocator.free(src.store_name);
            allocator.free(src.store_type);
            allocator.free(src.pst_path);
            if (err == error.Cancelled) break;
            continue;
        };

        try sources.append(allocator, .{ .source = src, .folder_selection = folder_sel });
    }

    if (sources.items.len == 0) {
        ui.failAbort("No hay origenes configurados.");
        return;
    }

    // Step 2: Select destination (common for all)
    const dest = try dest_selector.selectDestination(allocator, null);
    defer {
        allocator.free(dest.store_id);
        allocator.free(dest.store_name);
        allocator.free(dest.store_type);
        allocator.free(dest.pst_path);
    }

    // Step 3: Shared config
    var shared = shared_prompts.promptSharedConfig(allocator) catch {
        ui.failAbort("Error en la configuracion");
        return;
    };
    defer shared.deinit(allocator);

    // Step 4: Execute each source
    for (sources.items, 0..) |item, idx| {
        ui.clearScreen();
        ui.printSectionTitle("Progreso del Lote");
        std.debug.print("  \x1b[1;37mOrigen {d} de {d}:\x1b[0m {s}\n\n", .{ idx + 1, sources.items.len, item.source.store_name });

        const config = types.TransferConfig{
            .source_info = types.SourceInfo{
                .store_id = try allocator.dupe(u8, item.source.store_id),
                .store_name = try allocator.dupe(u8, item.source.store_name),
                .store_type = try allocator.dupe(u8, item.source.store_type),
                .pst_path = try allocator.dupe(u8, item.source.pst_path),
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
            .folder_plan_path = try allocator.dupe(u8, item.folder_selection.folder_plan_path),
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

        transfer_runner.executeTransfer(allocator, config) catch {
            ui.printError("Error procesando este origen, continuando con el siguiente...");
            ui.waitForEnter();
            continue;
        };
    }

    ui.clearScreen();
    ui.printSectionTitle("Lote completado");
    std.debug.print("  \x1b[1;32mProceso completado para {d} origenes.\x1b[0m\n\n", .{sources.items.len});
    ui.waitForEnter();
}

fn scanFolders(allocator: std.mem.Allocator, identifier: []const u8, is_pst: bool) !import_types.FolderSelectionResult {
    _ = is_pst;
    return try scan_service.runScanWithSource(
        allocator,
        identifier,
        .quick,
        null,
        null,
        false,
    );
}
