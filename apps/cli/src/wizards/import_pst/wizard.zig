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

pub fn run(allocator: std.mem.Allocator) !void {
    // Step 1: Select PST file
    const pst_path = file_browser.selectPstFile(allocator) catch |err| {
        if (err == error.Cancelled) {
            std.debug.print("\n  \x1b[90mOperacion cancelada.\x1b[0m\n", .{});
            ui.waitForEnter();
            return;
        }
        ui.printError("Error seleccionando archivo PST");
        ui.waitForEnter();
        return;
    };
    defer allocator.free(pst_path);

    // Step 2: Select Outlook profile
    const profile_name = prompts.chooseOutlookProfile(allocator) catch {
        ui.printError("Error seleccionando perfil de Outlook");
        ui.waitForEnter();
        return;
    };
    defer if (profile_name) |p| allocator.free(p);

    // Step 3: Scan PST and choose folders
    const scan_mode = prompts.chooseScanMode(allocator) catch {
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

    const folder_selection = runScanAndSelectFolders(allocator, pst_path, scan_mode, scan_filter_year, profile_name) catch |err| {
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
    }

    // Step 4: Action (Copy or Move)
    const action = prompts.chooseImportAction(pst_path);

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

    // Step 9: Select target store
    const selected_store = store_selector.selectTargetStore(allocator, profile_name) catch {
        ui.printError("Error seleccionando buzon destino");
        ui.waitForEnter();
        return;
    };
    defer allocator.free(selected_store.store_id);
    defer allocator.free(selected_store.display_name);
    defer allocator.free(selected_store.store_type);

    // Build config
    const config = types.ImportConfig{
        .pst_path = pst_path,
        .target_store_id = selected_store.store_id,
        .target_store_name = selected_store.display_name,
        .target_store_type = selected_store.store_type,
        .action = action,
        .skip_duplicates = skip_duplicates,
        .deep_duplicate_check = deep_duplicate_check,
        .filter_year = filter_year,
        .filter_months = filter_months,
        .folder_plan_path = folder_selection.folder_plan_path,
        .adaptive_throttling = adaptive_throttling,
        .profile_name = profile_name,
    };

    // Show summary and confirm
    ui.clearScreen();
    ui.printSectionTitle("Resumen de importacion");
    std.debug.print("  \x1b[1;37mArchivo PST:\x1b[0m   {s}\n", .{config.pst_path});
    const profile_display = if (config.profile_name) |p|
        if (p.len > 0) p else "Perfil predeterminado"
    else
        "Perfil predeterminado";
    std.debug.print("  \x1b[1;37mPerfil Outlook:\x1b[0m {s}\n", .{profile_display});
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

    // Execute import
    try executor.executeImport(allocator, config);
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
    if (scan_mode == .deep) {
        try args.append(allocator, "-IncludeSize");
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

    return .{
        .folder_plan_path = plan_path,
        .selected_count = selected_count,
        .total_count = folders.items.len,
        .scan_mode = scan_mode,
    };
}
