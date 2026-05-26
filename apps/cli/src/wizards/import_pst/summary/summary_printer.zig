const std = @import("std");
const ui = @import("../../../ui.zig");
const types = @import("../types.zig");
const utils = @import("../utils.zig");
const result_parser = @import("../execution/result_parser.zig");
const shared_config_mod = @import("../flow/shared_config.zig");

const ParsedImportOutput = result_parser.ParsedImportOutput;
const SharedConfig = shared_config_mod.SharedConfig;

fn actionDisplayName(action: []const u8) []const u8 {
    if (std.mem.eql(u8, action, "Move")) return "Mover";
    if (std.mem.eql(u8, action, "Copy")) return "Copiar";
    return action;
}

fn printRoutingMappingsSummary(criterion: types.RoutingCriterion, mappings: ?[]const types.TargetStoreMapping) void {
    const items = mappings orelse return;
    for (items) |m| {
        if (m.store_id.len == 0) continue;
        if (criterion == .by_year) {
            std.debug.print("    - Ano {d:4} => {s}\n", .{ m.year, m.store_name });
        } else {
            std.debug.print("    - {d:4}-{d:0>2} => {s}\n", .{ m.year, m.month.?, m.store_name });
        }
    }
}

pub fn printImportSummary(config: types.ImportConfig, size_str: []const u8, date_time_str: []const u8) void {
    std.debug.print("  \x1b[1;30m======================================================================\x1b[0m\n", .{});
    std.debug.print("  \x1b[1;37mPST de origen:\x1b[0m     {s} ({s})\n", .{ config.pst_path, size_str });
    std.debug.print("  \x1b[1;37mPerfil Outlook:\x1b[0m    {s}\n", .{utils.profileDisplayName(config.profile_name)});

    if (config.routing_criterion) |criterion| {
        std.debug.print("  \x1b[1;37mEnrutamiento:\x1b[0m      Multibuzon (agrupado por {s})\n", .{if (criterion == .by_year) "Anos" else "Meses"});
        printRoutingMappingsSummary(criterion, config.routing_mappings);
    } else {
        std.debug.print("  \x1b[1;37mBuzon de destino:\x1b[0m  {s}\n", .{config.target_store_name});
    }

    std.debug.print("  \x1b[1;37mTipo de buzon:\x1b[0m     {s}\n", .{utils.storeTypeDisplayName(config.target_store_type)});
    std.debug.print("  \x1b[1;37mAccion:\x1b[0m          {s}\n", .{actionDisplayName(config.action)});
    std.debug.print("  \x1b[1;37mInicio proceso:\x1b[0m    {s}\n", .{date_time_str});

    if (config.filter_year) |y| {
        std.debug.print("  \x1b[1;37mFiltro de anios:\x1b[0m  {s}\n", .{y});
    } else {
        std.debug.print("  \x1b[1;37mFiltro de anios:\x1b[0m  Todos los anios\n", .{});
    }
    std.debug.print("  \x1b[1;30m======================================================================\x1b[0m\n\n", .{});
}

pub fn printImportResult(parsed: *const ParsedImportOutput, elapsed_total: i64) void {
    std.debug.print("\n", .{});
    ui.printSectionTitle("Resultado de importacion");

    const json = parsed.result_line orelse "";
    if (json.len > 0) {
        const copied = utils.extractNumber(json, "copied") orelse parsed.last_copied;
        const moved = utils.extractNumber(json, "moved") orelse parsed.last_moved;
        const skipped = utils.extractNumber(json, "skipped") orelse parsed.last_skipped;
        const failed = utils.extractNumber(json, "failed") orelse parsed.last_failed;
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
        std.debug.print("  \x1b[1;32mCopiados:\x1b[0m    {d}\n", .{parsed.last_copied});
        std.debug.print("  \x1b[1;32mMovidos:\x1b[0m     {d}\n", .{parsed.last_moved});
        std.debug.print("  \x1b[1;33mOmitidos:\x1b[0m    {d}\n", .{parsed.last_skipped});
        std.debug.print("  \x1b[1;31mFallidos:\x1b[0m    {d}\n", .{parsed.last_failed});
        std.debug.print("  \x1b[90mTiempo:\x1b[0m      {d}s\n", .{@divTrunc(elapsed_total, 1000)});
    }
}

pub fn printSingleImportSummary(config: types.ImportConfig, shared: *const SharedConfig, folder_count: usize, folder_total: usize) void {
    ui.clearScreen();
    ui.printSectionTitle("Resumen de importacion");
    std.debug.print("  \x1b[1;37mArchivo PST:\x1b[0m   {s}\n", .{config.pst_path});
    std.debug.print("  \x1b[1;37mPerfil Outlook:\x1b[0m {s}\n", .{utils.profileDisplayName(config.profile_name)});

    if (config.routing_criterion) |criterion| {
        std.debug.print("  \x1b[1;37mEnrutamiento:\x1b[0m   {s}\n", .{utils.routingCriterionDisplay(criterion)});
        std.debug.print("  \x1b[1;37mBuzones:\x1b[0m       {d} asignados\n", .{utils.countAssignedMappings(config.routing_mappings)});
        printRoutingMappingsSummary(criterion, config.routing_mappings);
    } else {
        std.debug.print("  \x1b[1;37mBuzon destino:\x1b[0m {s}\n", .{config.target_store_name});
        std.debug.print("  \x1b[1;37mTipo de buzon:\x1b[0m {s}\n", .{utils.storeTypeDisplayName(config.target_store_type)});
    }

    std.debug.print("  \x1b[1;37mAccion:\x1b[0m        {s}\n", .{config.action});
    std.debug.print("  \x1b[1;37mEscaneo:\x1b[0m       {s}\n", .{if (shared.scan_mode == .deep) "Profundo" else "Rapido"});
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
    std.debug.print("  \x1b[1;37mPlan carpetas:\x1b[0m {s}\n", .{config.folder_plan_path});
    std.debug.print("  \x1b[1;37mThrottling:\x1b[0m    {s}\n", .{if (config.adaptive_throttling) "Adaptativo" else "Fijo"});
}

pub fn printBatchSummary(shared: *const SharedConfig, batch_configs: anytype) void {
    ui.clearScreen();
    ui.printSectionTitle("Resumen de lote");
    std.debug.print("  \x1b[1;37mPSTs configurados:\x1b[0m {d}\n", .{batch_configs.len});
    for (batch_configs, 0..) |cfg, idx| {
        std.debug.print("\n  \x1b[1;37m[{d}]\x1b[0m {s}\n", .{ idx + 1, cfg.pst_path });
        std.debug.print("      Buzon destino: {s}\n", .{cfg.target_store_name});
        std.debug.print("      Carpetas: {d}/{d}\n", .{ cfg.selected_count, cfg.total_count });
        std.debug.print("      Mapeos routing: {d}\n", .{utils.countAssignedMappings(cfg.routing_mappings)});
    }
    std.debug.print("  \x1b[1;37mEscaneo:\x1b[0m         {s}\n", .{if (shared.scan_mode == .deep) "Profundo" else "Rapido"});
    std.debug.print("  \x1b[1;37mAccion:\x1b[0m          {s}\n", .{shared.action});
    std.debug.print("  \x1b[1;37mDuplicados:\x1b[0m      {s}\n", .{if (shared.skip_duplicates) "Saltar" else "No saltar"});
    std.debug.print("  \x1b[1;37mThrottling:\x1b[0m      {s}\n", .{if (shared.adaptive_throttling) "Adaptativo" else "Fijo"});
    if (shared.routing_criterion) |criterion| {
        std.debug.print("  \x1b[1;37mEnrutamiento:\x1b[0m     {s}\n", .{if (criterion == .by_year) "Multibuzon por Anos" else "Multibuzon por Meses"});
    }
}

pub fn printBatchProgressHeader(index: usize, total: usize, pst_path: []const u8, target_store_name: []const u8) void {
    ui.clearScreen();
    ui.printSectionTitle("Progreso del Lote");
    std.debug.print("  \x1b[1;37mPST {d} de {d}:\x1b[0m {s}\n", .{ index, total, pst_path });
    if (target_store_name.len > 0) {
        std.debug.print("  \x1b[1;37mDestino:\x1b[0m {s}\n\n", .{target_store_name});
    } else {
        std.debug.print("\n", .{});
    }
}
