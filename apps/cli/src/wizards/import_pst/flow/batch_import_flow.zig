const std = @import("std");
const ui = @import("../../../ui.zig");
const types = @import("../types.zig");
const executor = @import("../executor.zig");
const import_config_factory = @import("import_config_factory.zig");
const summary_printer = @import("../summary/summary_printer.zig");
const single_import_flow = @import("single_import_flow.zig");
const shared_prompts_flow = @import("shared_prompts_flow.zig");
const batch_pst_config = @import("batch_pst_config.zig");
const shared_config_mod = @import("shared_config.zig");

const SharedConfig = shared_config_mod.SharedConfig;
const BatchPstConfig = batch_pst_config.BatchPstConfig;

fn collectBatchConfigs(
    allocator: std.mem.Allocator,
    pst_paths: [][]const u8,
    shared: *const SharedConfig,
) !std.ArrayListUnmanaged(BatchPstConfig) {
    var batch_configs = std.ArrayListUnmanaged(BatchPstConfig){};
    errdefer {
        for (batch_configs.items) |*cfg| cfg.deinit(allocator);
        batch_configs.deinit(allocator);
    }

    for (pst_paths, 0..) |current_pst, idx| {
        summary_printer.printBatchProgressHeader(idx + 1, pst_paths.len, current_pst, "");
        std.debug.print("  \x1b[90mSe escaneara este PST para mostrar anios/meses y definir destino(s).\x1b[0m\n\n", .{});

        var cfg = single_import_flow.configureOnePst(allocator, current_pst, shared) catch |err| {
            if (err == error.Cancelled) {
                std.debug.print("  \x1b[33mPST omitido por cancelacion.\x1b[0m\n", .{});
                ui.waitForEnter();
                continue;
            }
            ui.failAbort("Error escaneando PST o seleccionando carpetas para este archivo");
            continue;
        };

        batch_configs.append(allocator, cfg) catch {
            cfg.deinit(allocator);
            return error.OutOfMemory;
        };
    }

    return batch_configs;
}

fn executeBatchConfigs(
    allocator: std.mem.Allocator,
    shared: *const SharedConfig,
    batch_configs: []const BatchPstConfig,
) void {
    for (batch_configs, 0..) |cfg, idx| {
        summary_printer.printBatchProgressHeader(idx + 1, batch_configs.len, cfg.pst_path, cfg.target_store_name);

        const config = import_config_factory.buildImportConfig(
            shared,
            cfg.pst_path,
            cfg.target_store_id,
            cfg.target_store_name,
            cfg.target_store_type,
            cfg.folder_plan_path,
            cfg.routing_mappings,
        );

        executor.executeImport(allocator, config) catch |err| {
            ui.printError("Error procesando este archivo PST, continuando con el siguiente...");
            std.debug.print("  Detalle error: {}\n", .{err});
            ui.waitForEnter();
            continue;
        };
    }
}

fn deinitBatchConfigs(allocator: std.mem.Allocator, batch_configs: *std.ArrayListUnmanaged(BatchPstConfig)) void {
    for (batch_configs.items) |*cfg| cfg.deinit(allocator);
    batch_configs.deinit(allocator);
}

pub fn runMultiPstBatch(allocator: std.mem.Allocator, pst_paths: [][]const u8) !void {
    const reference_pst_path = pst_paths[0];

    var shared = shared_prompts_flow.promptSharedConfig(allocator, reference_pst_path) catch |err| {
        if (err == error.Cancelled) {
            std.debug.print("\n  \x1b[90mOperacion cancelada.\x1b[0m\n", .{});
            ui.waitForEnter();
            return;
        }
        ui.failAbort("Error en la configuracion inicial");
        return;
    };
    defer shared.deinit(allocator);

    var batch_configs = collectBatchConfigs(allocator, pst_paths, &shared) catch |err| {
        if (err == error.OutOfMemory) return error.OutOfMemory;
        return err;
    };
    defer deinitBatchConfigs(allocator, &batch_configs);

    if (batch_configs.items.len == 0) {
        ui.failAbort("No hay PST configurados para ejecutar en el lote.");
        return;
    }

    summary_printer.printBatchSummary(&shared, batch_configs.items);

    std.debug.print("\n  \x1b[1;33mIniciar lote? (S/n) [S]:\x1b[0m ", .{});
    const confirm = ui.readYesNo(true);
    if (!confirm) {
        std.debug.print("\n  \x1b[90mImportacion cancelada.\x1b[0m\n", .{});
        ui.waitForEnter();
        return;
    }

    executeBatchConfigs(allocator, &shared, batch_configs.items);

    ui.clearScreen();
    ui.printSectionTitle("Lote completado");
    std.debug.print("  \x1b[1;32mProceso completado secuencialmente para {d} PSTs configurados.\x1b[0m\n\n", .{batch_configs.items.len});
    ui.waitForEnter();
}
