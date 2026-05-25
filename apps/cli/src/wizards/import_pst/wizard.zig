const std = @import("std");
const ui = @import("../../ui.zig");
const file_browser = @import("../../file_browser.zig");
const store_selector = @import("../../store_selector.zig");
const ps_runner = @import("../../ps_runner.zig");
const types = @import("types.zig");
const utils = @import("utils.zig");
const prompts = @import("prompts.zig");
const folder_selector = @import("folder_selector.zig");
const executor = @import("executor.zig");
const multi_pst_selector = @import("multi_pst_selector.zig");

pub fn run(allocator: std.mem.Allocator) !void {
    ui.clearScreen();
    ui.printSectionTitle("Tipo de Importacion");
    std.debug.print("  \x1b[33m[1]\x1b[0m Un unico PST hacia uno o multiples buzones\n", .{});
    std.debug.print("  \x1b[33m[2]\x1b[0m Multiples PSTs hacia uno o multiples buzones (Lote secuencial)\n\n", .{});
    std.debug.print("  \x1b[33mSeleccione el modo de operacion [1]:\x1b[0m ", .{});
    const mode_key = ui.readSingleKey() catch '1';
    const is_multi_pst = (mode_key == '2');

    var psts_result: ?multi_pst_selector.PstListResult = null;
    var single_pst_path: ?[]u8 = null;
    defer {
        if (psts_result) |*r| r.deinit();
        if (single_pst_path) |p| allocator.free(p);
    }

    if (is_multi_pst) {
        psts_result = multi_pst_selector.selectMultiplePstFiles(allocator) catch |err| {
            if (err == error.Cancelled) {
                std.debug.print("\n  \x1b[90mOperacion cancelada.\x1b[0m\n", .{});
                ui.waitForEnter();
                return;
            }
            ui.printError("Error seleccionando archivos PST");
            ui.waitForEnter();
            return;
        };
    } else {
        single_pst_path = file_browser.selectPstFile(allocator) catch |err| {
            if (err == error.Cancelled) {
                std.debug.print("\n  \x1b[90mOperacion cancelada.\x1b[0m\n", .{});
                ui.waitForEnter();
                return;
            }
            ui.printError("Error seleccionando archivo PST");
            ui.waitForEnter();
            return;
        };
    }

    const pst_paths: [][]const u8 = if (is_multi_pst) blk: {
        const slice = allocator.alloc([]const u8, psts_result.?.paths.len) catch return;
        for (psts_result.?.paths, 0..) |path, i| {
            slice[i] = path;
        }
        break :blk slice;
    } else blk: {
        const slice = allocator.alloc([]const u8, 1) catch return;
        slice[0] = single_pst_path.?;
        break :blk slice;
    };
    defer allocator.free(pst_paths);

    if (is_multi_pst) {
        try runMultiPstBatch(allocator, pst_paths);
        return;
    }

    // Usamos el primer PST para mostrar en ciertos diálogos o flujos interactivos si es necesario
    const reference_pst_path = pst_paths[0];

    // Step 2: Select Outlook profile
    const profile_name = prompts.chooseOutlookProfile(allocator) catch {
        ui.printError("Error seleccionando perfil de Outlook");
        ui.waitForEnter();
        return;
    };
    defer if (profile_name) |p| allocator.free(p);

    // NUEVO PASO: Preguntar por el Criterio de Enrutamiento antes de seleccionar el PST y escanear
    ui.clearScreen();
    ui.printSectionTitle("Enrutamiento de Correos");
    std.debug.print("  \x1b[90mDeseas enrutar los correos hacia multiples buzones segun sus fechas?\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mActivar Enrutamiento de Correos? (s/N) [N]:\x1b[0m ", .{});
    const enable_routing = ui.readYesNo(false);

    var routing_criterion: ?types.RoutingCriterion = null;
    if (enable_routing) {
        const routing_wizard = @import("routing_wizard.zig");
        routing_criterion = routing_wizard.selectRoutingCriterion() catch |err| {
            if (err == error.Cancelled) {
                std.debug.print("\n  \x1b[90mOperacion cancelada.\x1b[0m\n", .{});
                ui.waitForEnter();
                return;
            }
            ui.printError("Error seleccionando criterio de enrutamiento");
            ui.waitForEnter();
            return;
        };
    }

    // Step 3: Scan PST and choose folders
    // Si el enrutamiento está habilitado, forzamos un escaneo "profundo" (con -ExportStatistics) para tener las fechas de meses/años
    const scan_mode = if (enable_routing) .deep else prompts.chooseScanMode(allocator) catch {
        ui.printError("Error leyendo modo de escaneo");
        ui.waitForEnter();
        return;
    };

    const scan_filter_year = prompts.chooseScanYearFilter(allocator) catch {
        ui.printError("Error leyendo filtro de anio para escaneo");
        ui.waitForEnter();
        return;
    };
    defer if (scan_filter_year) |y| allocator.free(y);

    const folder_selection = runScanAndSelectFolders(allocator, reference_pst_path, scan_mode, scan_filter_year, profile_name, enable_routing) catch |err| {
        if (err == error.Cancelled) {
            std.debug.print("\n  \x1b[90mOperacion cancelada.\x1b[0m\n", .{});
            ui.waitForEnter();
            return;
        }
        ui.printError("Error escaneando PST o seleccionando carpetas");
        ui.waitForEnter();
        return;
    };
    defer {
        utils.cleanupTempFile(folder_selection.folder_plan_path);
        allocator.free(folder_selection.folder_plan_path);
        utils.cleanupTempFile(folder_selection.scan_export_path);
        allocator.free(folder_selection.scan_export_path);
    }

    // Step 4: Action (Copy or Move)
    const action = prompts.chooseImportAction(reference_pst_path);

    // Step 5: Skip duplicates?
    ui.clearScreen();
    ui.printSectionTitle("Duplicados");
    std.debug.print("  \x1b[90mSaltar items duplicados detectados en el buzon destino?\x1b[0m\n", .{});
    std.debug.print("  \x1b[90mUsa Message-ID, SearchKey o clave compuesta para detectar.\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mSaltar duplicados? (S/n) [S]:\x1b[0m ", .{});
    const skip_duplicates = ui.readYesNo(true);

    // Step 5b: Deep duplicate check?
    var deep_duplicate_check = false;
    if (skip_duplicates) {
        std.debug.print("\n  \x1b[90mRevision profunda: tambien indexa subcarpetas del destino.\x1b[0m\n", .{});
        std.debug.print("  \x1b[90m(Mas lento, pero detecta duplicados movidos manualmente)\x1b[0m\n\n", .{});
        std.debug.print("  \x1b[33mRevision profunda? (s/N) [N]:\x1b[0m ", .{});
        deep_duplicate_check = ui.readYesNo(false);
    }

    // Step 6: Filter by year?
    const filter_year = promptOptionalYearFilter(allocator);
    defer if (filter_year) |y| allocator.free(y);

    // Step 7: Filter by months?
    const filter_months = promptOptionalMonthsFilter(allocator);
    defer if (filter_months) |m| allocator.free(m);

    // Step 8: Adaptive throttling?
    ui.clearScreen();
    ui.printSectionTitle("Rendimiento");
    std.debug.print("  \x1b[90mThrottling adaptativo: ajusta la velocidad automaticamente\x1b[0m\n", .{});
    std.debug.print("  \x1b[90mcuando Exchange limita las peticiones. Recomendado para\x1b[0m\n", .{});
    std.debug.print("  \x1b[90mExchange Online / Office 365.\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mActivar throttling adaptativo? (S/n) [S]:\x1b[0m ", .{});
    const adaptive_throttling = ui.readYesNo(true);

    // NUEVO PASO: Configurar Enrutamiento de los Correos (Mapeo de Buzones) si está activo
    var routing_mappings: ?[]const types.TargetStoreMapping = null;
    var selected_store_id: []const u8 = "";
    var selected_store_name: []const u8 = "";
    var selected_store_type: []const u8 = "";

    if (enable_routing) {
        const routing_wizard = @import("routing_wizard.zig");
        routing_mappings = routing_wizard.configureMappings(allocator, routing_criterion.?, folder_selection.scan_export_path, profile_name) catch |err| {
            if (err == error.Cancelled) {
                std.debug.print("\n  \x1b[90mOperacion cancelada.\x1b[0m\n", .{});
                ui.waitForEnter();
                return;
            }
            ui.printError("Error configurando el enrutamiento de buzones");
            ui.waitForEnter();
            return;
        };
        // Para compatibilidad con la estructura, usamos el primer buzón asignado como principal (o vacío)
        for (routing_mappings.?) |m| {
            if (m.store_id.len > 0) {
                selected_store_id = try allocator.dupe(u8, m.store_id);
                selected_store_name = try allocator.dupe(u8, m.store_name);
                selected_store_type = try allocator.dupe(u8, m.store_type);
                break;
            }
        }
    } else {
        // Step 9: Select target store (Buzón único de destino tradicional)
        const selected_store = store_selector.selectTargetStore(allocator, profile_name) catch {
            ui.printError("Error seleccionando buzon destino");
            ui.waitForEnter();
            return;
        };
        selected_store_id = selected_store.store_id;
        selected_store_name = selected_store.display_name;
        selected_store_type = selected_store.store_type;
    }
    defer allocator.free(selected_store_id);
    defer allocator.free(selected_store_name);
    defer allocator.free(selected_store_type);

    defer if (routing_mappings) |mappings| {
        for (mappings) |m| {
            allocator.free(m.store_id);
            allocator.free(m.store_name);
            allocator.free(m.store_type);
        }
        allocator.free(mappings);
    };

    // Build config
    const config = types.ImportConfig{
        .pst_path = reference_pst_path,
        .target_store_id = selected_store_id,
        .target_store_name = selected_store_name,
        .target_store_type = selected_store_type,
        .action = action,
        .skip_duplicates = skip_duplicates,
        .deep_duplicate_check = deep_duplicate_check,
        .filter_year = filter_year,
        .filter_months = filter_months,
        .folder_plan_path = folder_selection.folder_plan_path,
        .adaptive_throttling = adaptive_throttling,
        .profile_name = profile_name,
        .routing_criterion = routing_criterion,
        .routing_mappings = routing_mappings,
    };

    // Show summary and confirm
    ui.clearScreen();
    ui.printSectionTitle("Resumen de importacion");
    if (pst_paths.len > 1) {
        std.debug.print("  \x1b[1;37mArchivos PST ({d}):\x1b[0m\n", .{pst_paths.len});
        for (pst_paths) |p| {
            std.debug.print("    - {s}\n", .{p});
        }
    } else {
        std.debug.print("  \x1b[1;37mArchivo PST:\x1b[0m   {s}\n", .{config.pst_path});
    }
    const profile_display = if (config.profile_name) |p|
        if (p.len > 0) p else "Perfil predeterminado"
    else
        "Perfil predeterminado";
    std.debug.print("  \x1b[1;37mPerfil Outlook:\x1b[0m {s}\n", .{profile_display});

    if (config.routing_criterion) |criterion| {
        std.debug.print("  \x1b[1;37mEnrutamiento:\x1b[0m   {s}\n", .{if (criterion == .by_year) "Múltiples buzones agrupados por Años" else "Múltiples buzones agrupados por Meses"});
        var mapped_count: usize = 0;
        if (config.routing_mappings) |mappings| {
            for (mappings) |m| {
                if (m.store_id.len > 0) {
                    mapped_count += 1;
                    if (criterion == .by_year) {
                        std.debug.print("    - Año {d:4} => {s}\n", .{ m.year, m.store_name });
                    } else {
                        std.debug.print("    - {d:4}-{d:0>2} => {s}\n", .{ m.year, m.month.?, m.store_name });
                    }
                }
            }
        }
        std.debug.print("  \x1b[1;37mBuzones:\x1b[0m       {d} asignados\n", .{mapped_count});
    } else {
        std.debug.print("  \x1b[1;37mBuzon destino:\x1b[0m {s}\n", .{config.target_store_name});
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
        std.debug.print("  \x1b[1;37mTipo de buzon:\x1b[0m {s}\n", .{store_type_friendly});
    }
    std.debug.print("  \x1b[1;37mAccion:\x1b[0m        {s}\n", .{config.action});
    std.debug.print("  \x1b[1;37mEscaneo:\x1b[0m       {s}\n", .{if (folder_selection.scan_mode == .deep) "Profundo" else "Rapido"});
    std.debug.print("  \x1b[1;37mCarpetas:\x1b[0m      {d}/{d} seleccionadas\n", .{ folder_selection.selected_count, folder_selection.total_count });
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

    std.debug.print("\n  \x1b[1;33mIniciar importacion? (S/n) [S]:\x1b[0m ", .{});
    const confirm = ui.readYesNo(true);

    if (!confirm) {
        std.debug.print("\n  \x1b[90mImportacion cancelada.\x1b[0m\n", .{});
        ui.waitForEnter();
        return;
    }

    try executor.executeImport(allocator, config);
}

fn runMultiPstBatch(allocator: std.mem.Allocator, pst_paths: [][]const u8) !void {
    const BatchPstConfig = struct {
        pst_path: []const u8,
        folder_plan_path: []const u8,
        scan_export_path: []const u8,
        selected_count: usize,
        total_count: usize,
        target_store_id: []const u8,
        target_store_name: []const u8,
        target_store_type: []const u8,
        routing_mappings: ?[]const types.TargetStoreMapping,
    };

    const reference_pst_path = pst_paths[0];

    const profile_name = prompts.chooseOutlookProfile(allocator) catch {
        ui.printError("Error seleccionando perfil de Outlook");
        ui.waitForEnter();
        return;
    };
    defer if (profile_name) |p| allocator.free(p);

    ui.clearScreen();
    ui.printSectionTitle("Enrutamiento de Correos");
    std.debug.print("  \x1b[90mDeseas enrutar los correos hacia multiples buzones segun sus fechas?\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mActivar Enrutamiento de Correos? (s/N) [N]:\x1b[0m ", .{});
    const enable_routing = ui.readYesNo(false);

    var routing_criterion: ?types.RoutingCriterion = null;
    if (enable_routing) {
        const routing_wizard = @import("routing_wizard.zig");
        routing_criterion = routing_wizard.selectRoutingCriterion() catch |err| {
            if (err == error.Cancelled) {
                std.debug.print("\n  \x1b[90mOperacion cancelada.\x1b[0m\n", .{});
                ui.waitForEnter();
                return;
            }
            ui.printError("Error seleccionando criterio de enrutamiento");
            ui.waitForEnter();
            return;
        };
    }

    const scan_mode = if (enable_routing) .deep else prompts.chooseScanMode(allocator) catch {
        ui.printError("Error leyendo modo de escaneo");
        ui.waitForEnter();
        return;
    };

    const scan_filter_year = prompts.chooseScanYearFilter(allocator) catch {
        ui.printError("Error leyendo filtro de anio para escaneo");
        ui.waitForEnter();
        return;
    };
    defer if (scan_filter_year) |y| allocator.free(y);

    const action = prompts.chooseImportAction(reference_pst_path);

    ui.clearScreen();
    ui.printSectionTitle("Duplicados");
    std.debug.print("  \x1b[90mSaltar items duplicados detectados en el buzon destino?\x1b[0m\n", .{});
    std.debug.print("  \x1b[90mUsa Message-ID, SearchKey o clave compuesta para detectar.\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mSaltar duplicados? (S/n) [S]:\x1b[0m ", .{});
    const skip_duplicates = ui.readYesNo(true);

    var deep_duplicate_check = false;
    if (skip_duplicates) {
        std.debug.print("\n  \x1b[90mRevision profunda: tambien indexa subcarpetas del destino.\x1b[0m\n", .{});
        std.debug.print("  \x1b[90m(Mas lento, pero detecta duplicados movidos manualmente)\x1b[0m\n\n", .{});
        std.debug.print("  \x1b[33mRevision profunda? (s/N) [N]:\x1b[0m ", .{});
        deep_duplicate_check = ui.readYesNo(false);
    }

    const filter_year = promptOptionalYearFilter(allocator);
    defer if (filter_year) |y| allocator.free(y);

    const filter_months = promptOptionalMonthsFilter(allocator);
    defer if (filter_months) |m| allocator.free(m);

    ui.clearScreen();
    ui.printSectionTitle("Rendimiento");
    std.debug.print("  \x1b[90mThrottling adaptativo: ajusta la velocidad automaticamente\x1b[0m\n", .{});
    std.debug.print("  \x1b[90mcuando Exchange limita las peticiones. Recomendado para\x1b[0m\n", .{});
    std.debug.print("  \x1b[90mExchange Online / Office 365.\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mActivar throttling adaptativo? (S/n) [S]:\x1b[0m ", .{});
    const adaptive_throttling = ui.readYesNo(true);

    var batch_configs = std.ArrayListUnmanaged(BatchPstConfig){};
    defer {
        for (batch_configs.items) |cfg| {
            utils.cleanupTempFile(cfg.folder_plan_path);
            allocator.free(cfg.folder_plan_path);
            utils.cleanupTempFile(cfg.scan_export_path);
            allocator.free(cfg.scan_export_path);
            allocator.free(cfg.target_store_id);
            allocator.free(cfg.target_store_name);
            allocator.free(cfg.target_store_type);
            if (cfg.routing_mappings) |mappings| {
                for (mappings) |m| {
                    allocator.free(m.store_id);
                    allocator.free(m.store_name);
                    allocator.free(m.store_type);
                }
                allocator.free(mappings);
            }
        }
        batch_configs.deinit(allocator);
    }

    for (pst_paths, 0..) |current_pst, idx| {
        ui.clearScreen();
        ui.printSectionTitle("Configurar PST del Lote");
        std.debug.print("  \x1b[1;37mPST {d} de {d}:\x1b[0m {s}\n\n", .{ idx + 1, pst_paths.len, current_pst });
        std.debug.print("  \x1b[90mSe escaneara este PST para mostrar anios/meses y definir destino(s).\x1b[0m\n\n", .{});

        const folder_selection = runScanAndSelectFolders(allocator, current_pst, scan_mode, scan_filter_year, profile_name, enable_routing) catch |err| {
            if (err == error.Cancelled) {
                std.debug.print("  \x1b[33mPST omitido por cancelacion de seleccion de carpetas.\x1b[0m\n", .{});
                ui.waitForEnter();
                continue;
            }
            ui.printError("Error escaneando PST o seleccionando carpetas para este archivo");
            ui.waitForEnter();
            continue;
        };

        var routing_mappings: ?[]const types.TargetStoreMapping = null;
        var selected_store_id: ?[]const u8 = null;
        var selected_store_name: ?[]const u8 = null;
        var selected_store_type: ?[]const u8 = null;

        if (enable_routing) {
            const routing_wizard = @import("routing_wizard.zig");
            routing_mappings = routing_wizard.configureMappings(allocator, routing_criterion.?, folder_selection.scan_export_path, profile_name) catch |err| {
                utils.cleanupTempFile(folder_selection.folder_plan_path);
                allocator.free(folder_selection.folder_plan_path);
                utils.cleanupTempFile(folder_selection.scan_export_path);
                allocator.free(folder_selection.scan_export_path);
                if (err == error.Cancelled) {
                    std.debug.print("  \x1b[33mPST omitido por cancelacion en configuracion de enrutamiento.\x1b[0m\n", .{});
                    ui.waitForEnter();
                    continue;
                }
                ui.printError("Error configurando el enrutamiento para este PST");
                ui.waitForEnter();
                continue;
            };

            for (routing_mappings.?) |m| {
                if (m.store_id.len > 0) {
                    selected_store_id = try allocator.dupe(u8, m.store_id);
                    selected_store_name = try allocator.dupe(u8, m.store_name);
                    selected_store_type = try allocator.dupe(u8, m.store_type);
                    break;
                }
            }

            if (selected_store_id == null or selected_store_name == null or selected_store_type == null) {
                if (routing_mappings) |mappings| {
                    for (mappings) |m| {
                        allocator.free(m.store_id);
                        allocator.free(m.store_name);
                        allocator.free(m.store_type);
                    }
                    allocator.free(mappings);
                }
                utils.cleanupTempFile(folder_selection.folder_plan_path);
                allocator.free(folder_selection.folder_plan_path);
                utils.cleanupTempFile(folder_selection.scan_export_path);
                allocator.free(folder_selection.scan_export_path);
                ui.printError("No se asigno ningun buzon en el mapeo de enrutamiento para este PST");
                ui.waitForEnter();
                continue;
            }
        } else {
            ui.clearScreen();
            ui.printSectionTitle("Destino para PST");
            std.debug.print("  \x1b[1;37mPST {d} de {d}:\x1b[0m {s}\n", .{ idx + 1, pst_paths.len, current_pst });
            std.debug.print("  \x1b[90mCarpetas seleccionadas: {d}/{d}. Elige el buzon destino para este PST.\x1b[0m\n\n", .{ folder_selection.selected_count, folder_selection.total_count });

            const selected_store = store_selector.selectTargetStore(allocator, profile_name) catch {
                utils.cleanupTempFile(folder_selection.folder_plan_path);
                allocator.free(folder_selection.folder_plan_path);
                utils.cleanupTempFile(folder_selection.scan_export_path);
                allocator.free(folder_selection.scan_export_path);
                ui.printError("Error seleccionando buzon destino para este PST");
                ui.waitForEnter();
                continue;
            };
            selected_store_id = selected_store.store_id;
            selected_store_name = selected_store.display_name;
            selected_store_type = selected_store.store_type;
        }

        batch_configs.append(allocator, .{
            .pst_path = current_pst,
            .folder_plan_path = folder_selection.folder_plan_path,
            .scan_export_path = folder_selection.scan_export_path,
            .selected_count = folder_selection.selected_count,
            .total_count = folder_selection.total_count,
            .target_store_id = selected_store_id.?,
            .target_store_name = selected_store_name.?,
            .target_store_type = selected_store_type.?,
            .routing_mappings = routing_mappings,
        }) catch {
            if (routing_mappings) |mappings| {
                for (mappings) |m| {
                    allocator.free(m.store_id);
                    allocator.free(m.store_name);
                    allocator.free(m.store_type);
                }
                allocator.free(mappings);
            }
            allocator.free(selected_store_id.?);
            allocator.free(selected_store_name.?);
            allocator.free(selected_store_type.?);
            utils.cleanupTempFile(folder_selection.folder_plan_path);
            allocator.free(folder_selection.folder_plan_path);
            utils.cleanupTempFile(folder_selection.scan_export_path);
            allocator.free(folder_selection.scan_export_path);
            return error.OutOfMemory;
        };
    }

    if (batch_configs.items.len == 0) {
        ui.printError("No hay PST configurados para ejecutar en el lote.");
        ui.waitForEnter();
        return;
    }

    ui.clearScreen();
    ui.printSectionTitle("Resumen de lote");
    std.debug.print("  \x1b[1;37mPSTs configurados:\x1b[0m {d}\n", .{batch_configs.items.len});
    for (batch_configs.items, 0..) |cfg, idx| {
        std.debug.print("\n  \x1b[1;37m[{d}]\x1b[0m {s}\n", .{ idx + 1, cfg.pst_path });
        std.debug.print("      Buzon destino: {s}\n", .{cfg.target_store_name});
        std.debug.print("      Carpetas: {d}/{d}\n", .{ cfg.selected_count, cfg.total_count });
        if (enable_routing and cfg.routing_mappings != null) {
            var mapped_count: usize = 0;
            for (cfg.routing_mappings.?) |m| {
                if (m.store_id.len > 0) mapped_count += 1;
            }
            std.debug.print("      Mapeos routing: {d}\n", .{mapped_count});
        }
    }
    std.debug.print("  \x1b[1;37mEscaneo:\x1b[0m         {s}\n", .{if (scan_mode == .deep) "Profundo" else "Rapido"});
    std.debug.print("  \x1b[1;37mAccion:\x1b[0m          {s}\n", .{action});
    std.debug.print("  \x1b[1;37mDuplicados:\x1b[0m      {s}\n", .{if (skip_duplicates) "Saltar" else "No saltar"});
    std.debug.print("  \x1b[1;37mThrottling:\x1b[0m      {s}\n", .{if (adaptive_throttling) "Adaptativo" else "Fijo"});
    if (routing_criterion) |criterion| {
        std.debug.print("  \x1b[1;37mEnrutamiento:\x1b[0m     {s}\n", .{if (criterion == .by_year) "Multibuzon por Anos" else "Multibuzon por Meses"});
    }

    std.debug.print("\n  \x1b[1;33mIniciar lote? (S/n) [S]:\x1b[0m ", .{});
    const confirm = ui.readYesNo(true);
    if (!confirm) {
        std.debug.print("\n  \x1b[90mImportacion cancelada.\x1b[0m\n", .{});
        ui.waitForEnter();
        return;
    }

    for (batch_configs.items, 0..) |cfg, idx| {
        ui.clearScreen();
        ui.printSectionTitle("Progreso del Lote");
        std.debug.print("  \x1b[1;37mPST {d} de {d}:\x1b[0m {s}\n", .{ idx + 1, batch_configs.items.len, cfg.pst_path });
        std.debug.print("  \x1b[1;37mDestino:\x1b[0m {s}\n\n", .{cfg.target_store_name});

        const config = types.ImportConfig{
            .pst_path = cfg.pst_path,
            .target_store_id = cfg.target_store_id,
            .target_store_name = cfg.target_store_name,
            .target_store_type = cfg.target_store_type,
            .action = action,
            .skip_duplicates = skip_duplicates,
            .deep_duplicate_check = deep_duplicate_check,
            .filter_year = filter_year,
            .filter_months = filter_months,
            .folder_plan_path = cfg.folder_plan_path,
            .adaptive_throttling = adaptive_throttling,
            .profile_name = profile_name,
            .routing_criterion = routing_criterion,
            .routing_mappings = cfg.routing_mappings,
        };

        executor.executeImport(allocator, config) catch |err| {
            ui.printError("Error procesando este archivo PST, continuando con el siguiente...");
            std.debug.print("  Detalle error: {}\n", .{err});
            ui.waitForEnter();
            continue;
        };
    }

    ui.clearScreen();
    ui.printSectionTitle("Lote completado");
    std.debug.print("  \x1b[1;32mProceso completado secuencialmente para {d} PSTs configurados.\x1b[0m\n\n", .{batch_configs.items.len});
    ui.waitForEnter();
}

fn promptOptionalYearFilter(allocator: std.mem.Allocator) ?[]const u8 {
    ui.clearScreen();
    ui.printSectionTitle("Filtros de fecha");
    std.debug.print("  \x1b[90mFiltrar por anio? Dejar vacio para importar todos.\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mAnio (ej: 2023) o Enter para todos:\x1b[0m ", .{});

    const year_input = ui.readLine(allocator) catch return null;
    if (year_input.len == 0) {
        allocator.free(year_input);
        return null;
    }

    const year_value = std.fmt.parseInt(i32, year_input, 10) catch {
        ui.printError("Anio invalido, se ignorara el filtro");
        allocator.free(year_input);
        ui.waitForEnter();
        return null;
    };

    if (year_value < 1900 or year_value > 9999) {
        ui.printError("Anio fuera de rango (1900-9999), se ignorara el filtro");
        allocator.free(year_input);
        ui.waitForEnter();
        return null;
    }

    return year_input;
}

fn promptOptionalMonthsFilter(allocator: std.mem.Allocator) ?[]const u8 {
    std.debug.print("\n  \x1b[90mFiltrar por meses? Separar con comas.\x1b[0m\n", .{});
    std.debug.print("  \x1b[90mAcepta: numeros (1-12), nombres (enero, feb, march)\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mMeses (ej: ene,feb,mar) o Enter para todos:\x1b[0m ", .{});

    const months_input = ui.readLine(allocator) catch return null;
    if (months_input.len == 0) {
        allocator.free(months_input);
        return null;
    }

    return months_input;
}

fn runScanAndSelectFolders(
    allocator: std.mem.Allocator,
    pst_path: []const u8,
    scan_mode: types.ScanMode,
    scan_filter_year: ?[]const u8,
    profile_name: ?[]const u8,
    enable_routing: bool,
) !types.FolderSelectionResult {
    ui.clearScreen();
    ui.printSectionTitle("Escanear PST");
    std.debug.print("  \x1b[90mEjecutando escaneo {s}...\x1b[0m\n", .{if (scan_mode == .deep) "profundo" else "rapido"});

    const scan_script_path = try ps_runner.writeEmbeddedScript(allocator, .scan_pst);
    defer ps_runner.cleanupScript(allocator, scan_script_path);

    const exported_scan_path = try utils.makeTempFilePath(allocator, "oo-scan-export", "json");
    defer allocator.free(exported_scan_path);

    var args = std.ArrayListUnmanaged([]const u8){};
    defer args.deinit(allocator);

    try args.appendSlice(allocator, &.{
        "-PstPath",
        pst_path,
        "-Json",
        "-Headless",
        "-ExportFolders",
        "-ExportFoldersPath",
        exported_scan_path,
    });
    if (scan_mode == .deep or enable_routing) {
        try args.append(allocator, "-IncludeSize");
    }
    // Si el enrutamiento está habilitado, forzamos -ExportStatistics en el escaneo para tener los meses
    if (enable_routing) {
        try args.append(allocator, "-ExportStatistics");
    }
    if (scan_filter_year) |y| {
        try args.append(allocator, "-FilterOnlyYear");
        try args.append(allocator, y);
    }
    if (profile_name) |p| {
        if (p.len > 0) {
            try args.append(allocator, "-ProfileName");
            try args.append(allocator, p);
        }
    }

    const output = ps_runner.runScript(allocator, scan_script_path, args.items) catch return error.ScanFailed;
    defer allocator.free(output);

    var folders = std.ArrayListUnmanaged(types.ScannedFolder){};
    defer {
        for (folders.items) |folder| {
            allocator.free(folder.path);
            allocator.free(folder.year_breakdown_display);
        }
        folders.deinit(allocator);
    }

    var line_iter = std.mem.splitSequence(u8, output, "\n");
    while (line_iter.next()) |raw_line| {
        const line = std.mem.trim(u8, raw_line, &[_]u8{ ' ', '\t', '\r', '\n' });
        if (line.len == 0) continue;
        if (std.mem.indexOf(u8, line, "\"type\":\"folder\"") == null) continue;

        const item_count = utils.extractNumber(line, "itemCount") orelse 0;
        if (item_count <= 0) continue;

        const path = utils.extractString(line, "path") orelse continue;
        const path_copy = try allocator.dupe(u8, path);
        errdefer allocator.free(path_copy);

        const size_bytes = utils.extractNumber(line, "sizeBytes");
        const year_breakdown_display = try utils.extractYearBreakdownDisplay(allocator, line);
        errdefer allocator.free(year_breakdown_display);

        try folders.append(allocator, .{
            .path = path_copy,
            .item_count = item_count,
            .size_bytes = size_bytes,
            .year_breakdown_display = year_breakdown_display,
        });
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
