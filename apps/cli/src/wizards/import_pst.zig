const std = @import("std");
const ui = @import("../ui.zig");
const file_browser = @import("../file_browser.zig");
const store_selector = @import("../store_selector.zig");
const ps_runner = @import("../ps_runner.zig");

const ScanMode = enum {
    quick,
    deep,
};

const ImportProgressState = struct {
    start_ms: i64,
    copied: i64,
    moved: i64,
    skipped: i64,
    failed: i64,
    size_bytes: i64,
    percent: u32,
    has_rendered_progress: bool,
};

const ScannedFolder = struct {
    path: []u8,
    item_count: i64,
    size_bytes: ?i64,
    year_breakdown_display: []u8,
};

const FolderTreeNode = struct {
    name: []u8,
    full_path: []u8,
    parent: ?usize,
    children: std.ArrayListUnmanaged(usize),
    expanded: bool,
    folder_index: ?usize,
};

const VisibleTreeRow = struct {
    node_index: usize,
    depth: usize,
};

const FolderTree = struct {
    nodes: std.ArrayListUnmanaged(FolderTreeNode),
    roots: std.ArrayListUnmanaged(usize),
};

const FolderSelectionResult = struct {
    folder_plan_path: []u8,
    selected_count: usize,
    total_count: usize,
    scan_mode: ScanMode,
};

const ImportConfig = struct {
    pst_path: []const u8,
    target_store_id: []const u8,
    target_store_name: []const u8,
    target_store_type: []const u8,
    action: []const u8,
    skip_duplicates: bool,
    deep_duplicate_check: bool,
    filter_year: ?[]const u8,
    filter_months: ?[]const u8,
    folder_plan_path: []const u8,
    adaptive_throttling: bool,
    profile_name: ?[]const u8,
};

fn chooseOutlookProfile(allocator: std.mem.Allocator) !?[]const u8 {
    var cursor: usize = 0;
    const labels = [_][]const u8{
        "Usar perfil predeterminado de Outlook",
        "Especificar nombre del perfil",
    };

    while (true) {
        ui.clearScreen();
        ui.printSectionTitle("Seleccionar perfil de Outlook");
        std.debug.print("  \x1b[90mSelecciona el perfil de Outlook que deseas usar para esta importacion.\x1b[0m\n", .{});
        std.debug.print("  \x1b[90m\xe2\x86\x91/\xe2\x86\x93 mover | Enter confirmar\x1b[0m\n\n", .{});

        for (labels, 0..) |label, idx| {
            const is_current = idx == cursor;
            if (is_current) std.debug.print("  \x1b[7m", .{});
            std.debug.print("  {s}\n", .{label});
            if (is_current) std.debug.print("  \x1b[0m", .{});
        }

        const key = ui.readSingleKey() catch continue;
        switch (key) {
            'w', 'W', 'k', 'K' => {
                if (cursor > 0) cursor -= 1;
            },
            's', 'S', 'j', 'J' => {
                if (cursor + 1 < labels.len) cursor += 1;
            },
            '\r', '\n' => {
                if (cursor == 0) {
                    return null;
                }

                while (true) {
                    ui.clearScreen();
                    ui.printSectionTitle("Especificar perfil de Outlook");
                    std.debug.print("  \x1b[90mEscribe el nombre del perfil de Outlook que deseas usar.\x1b[0m\n\n", .{});
                    std.debug.print("  \x1b[33mNombre del perfil:\x1b[0m ", .{});

                    const input = ui.readLine(allocator) catch continue;
                    if (input.len == 0) {
                        allocator.free(input);
                        continue;
                    }
                    return input;
                }
            },
            27 => {
                const seq1 = ui.readSingleKey() catch continue;
                if (seq1 != '[' and seq1 != 'O') continue;

                const seq2 = ui.readSingleKey() catch continue;
                switch (seq2) {
                    'A' => {
                        if (cursor > 0) cursor -= 1;
                    },
                    'B' => {
                        if (cursor + 1 < labels.len) cursor += 1;
                    },
                    else => {},
                }
            },
            0, 224 => {
                const ext = ui.readSingleKey() catch continue;
                switch (ext) {
                    72 => {
                        if (cursor > 0) cursor -= 1;
                    },
                    80 => {
                        if (cursor + 1 < labels.len) cursor += 1;
                    },
                    else => {},
                }
            },
            else => {},
        }
    }
}

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
    const profile_name = chooseOutlookProfile(allocator) catch {
        ui.printError("Error seleccionando perfil de Outlook");
        ui.waitForEnter();
        return;
    };
    defer if (profile_name) |p| allocator.free(p);

    // Step 3: Scan PST and choose folders
    const scan_mode = chooseScanMode(allocator) catch {
        ui.printError("Error leyendo modo de escaneo");
        ui.waitForEnter();
        return;
    };

    const scan_filter_year = chooseScanYearFilter(allocator) catch {
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
        cleanupTempFile(folder_selection.folder_plan_path);
        allocator.free(folder_selection.folder_plan_path);
    }

    // Step 4: Action (Copy or Move)
    const action = chooseImportAction(pst_path);

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
    ui.clearScreen();
    ui.printSectionTitle("Filtros de fecha");
    std.debug.print("  \x1b[90mFiltrar por anio? Dejar vacio para importar todos.\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mAnio (ej: 2023) o Enter para todos:\x1b[0m ", .{});

    const year_input = ui.readLine(allocator) catch "";
    var filter_year: ?[]const u8 = null;
    if (year_input.len > 0) {
        // Validate it's a number
        _ = std.fmt.parseInt(i32, year_input, 10) catch {
            ui.printError("Anio invalido, se ignorara el filtro");
            allocator.free(year_input);
            ui.waitForEnter();
        };
        if (year_input.len > 0) {
            // Re-check since we may have freed it
            _ = std.fmt.parseInt(i32, year_input, 10) catch {
                filter_year = null;
            };
            filter_year = year_input;
        }
    }
    defer if (filter_year) |y| allocator.free(y);

    // Step 7: Filter by months?
    var filter_months: ?[]const u8 = null;
    std.debug.print("\n  \x1b[90mFiltrar por meses? Separar con comas.\x1b[0m\n", .{});
    std.debug.print("  \x1b[90mAcepta: numeros (1-12), nombres (enero, feb, march)\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mMeses (ej: ene,feb,mar) o Enter para todos:\x1b[0m ", .{});

    const months_input = ui.readLine(allocator) catch "";
    if (months_input.len > 0) {
        filter_months = months_input;
    } else {
        if (months_input.len == 0) {} // readLine returns empty string
    }
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
    const config = ImportConfig{
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
    try executeImport(allocator, config);
}

fn onImportScriptLine(ctx: *anyopaque, line: []const u8) void {
    var state: *ImportProgressState = @ptrCast(@alignCast(ctx));
    if (line.len == 0) return;
    if (std.mem.indexOf(u8, line, "\"type\":\"progress\"") == null) return;

    if (extractNumber(line, "copied")) |v| state.copied = v;
    if (extractNumber(line, "moved")) |v| state.moved = v;
    if (extractNumber(line, "skipped")) |v| state.skipped = v;
    if (extractNumber(line, "failed")) |v| state.failed = v;
    if (extractNumber(line, "sizeBytes")) |v| state.size_bytes = v;
    if (extractNumber(line, "percent")) |v| {
        if (v <= 0) {
            state.percent = 0;
        } else if (v >= 100) {
            state.percent = 100;
        } else {
            state.percent = @intCast(v);
        }
    }

    const total_script_processed = state.copied + state.moved + state.skipped + state.failed;
    const percent_effective: u32 = if (state.percent > 0)
        if (state.percent >= 100) 100 else state.percent
    else
        @as(u32, if (total_script_processed > 0) 1 else 0);
    var total_estimated = total_script_processed;
    if (percent_effective > 0) {
        var est = @divTrunc(total_script_processed * 100, @as(i64, @intCast(percent_effective)));
        if (est < total_script_processed) est = total_script_processed;
        total_estimated = est;
    }
    const remaining = @max(total_estimated - total_script_processed, 0);
    const elapsed_ms = std.time.milliTimestamp() - state.start_ms;

    if (state.has_rendered_progress) {
        std.debug.print("\x1b[2F", .{});
    } else {
        state.has_rendered_progress = true;
    }

    const current_status = extractString(line, "status") orelse "";
    std.debug.print("\r\x1b[2K  \x1b[90mCarpeta:\x1b[0m {s}\n", .{current_status});

    ui.printProgressBar(percent_effective, "Importando");

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

    std.debug.print("\n\x1b[2K  \x1b[90mProcesados:\x1b[0m {d} ({s}) \x1b[90m| Copiados:\x1b[0m {d} \x1b[90m| Omitidos:\x1b[0m {d} \x1b[90m| Restantes(est):\x1b[0m {d} \x1b[90m| Transcurrido:\x1b[0m {d:0>2}:{d:0>2}:{d:0>2}", .{
        total_script_processed,
        size_str,
        state.copied + state.moved,
        state.skipped + state.failed,
        remaining,
        elapsed_parts.hours,
        elapsed_parts.minutes,
        elapsed_parts.seconds,
    });

    if (percent_effective > 0 and percent_effective < 100) {
        const total_est_ms = @divTrunc(elapsed_ms * 100, @as(i64, @intCast(percent_effective)));
        const eta_ms = @max(total_est_ms - elapsed_ms, 0);
        const eta_s = @divTrunc(eta_ms, 1000);
        const eta_parts = ui.secondsToHms(eta_s);
        std.debug.print(
            " \x1b[90m| ETA:\x1b[0m {d:0>2}:{d:0>2}:{d:0>2}",
            .{ eta_parts.hours, eta_parts.minutes, eta_parts.seconds },
        );
    }
    std.debug.print("\r", .{});
}

fn chooseScanMode(allocator: std.mem.Allocator) !ScanMode {
    _ = allocator;
    var cursor: usize = 0;

    while (true) {
        ui.clearScreen();
        ui.printSectionTitle("Escanear PST");
        std.debug.print("  \x1b[90mSelecciona el tipo de escaneo antes de definir la accion.\x1b[0m\n", .{});
        std.debug.print("  \x1b[90m↑/↓ mover | Enter confirmar\x1b[0m\n\n", .{});

        const labels = [_][]const u8{
            "Escaneo Rapido  (carpetas / conteo de items)",
            "Escaneo Profundo (carpetas / conteo de items / tamano)",
        };

        for (labels, 0..) |label, idx| {
            const is_current = idx == cursor;
            if (is_current) std.debug.print("  \x1b[7m", .{});
            std.debug.print("  {s}\n", .{label});
            if (is_current) std.debug.print("  \x1b[0m", .{});
        }

        const key = ui.readSingleKey() catch continue;
        switch (key) {
            'w', 'W', 'k', 'K' => {
                if (cursor > 0) cursor -= 1;
            },
            's', 'S', 'j', 'J' => {
                if (cursor + 1 < labels.len) cursor += 1;
            },
            '1' => return .quick,
            '2' => return .deep,
            '\r', '\n' => return if (cursor == 0) .quick else .deep,
            27 => {
                const seq1 = ui.readSingleKey() catch continue;
                if (seq1 != '[' and seq1 != 'O') continue;

                const seq2 = ui.readSingleKey() catch continue;
                switch (seq2) {
                    'A' => {
                        if (cursor > 0) cursor -= 1;
                    },
                    'B' => {
                        if (cursor + 1 < labels.len) cursor += 1;
                    },
                    else => {},
                }
            },
            0, 224 => {
                const ext = ui.readSingleKey() catch continue;
                switch (ext) {
                    72 => {
                        if (cursor > 0) cursor -= 1;
                    },
                    80 => {
                        if (cursor + 1 < labels.len) cursor += 1;
                    },
                    else => {},
                }
            },
            else => {},
        }
    }
}

fn chooseImportAction(pst_path: []const u8) []const u8 {
    var cursor: usize = 0;
    const labels = [_][]const u8{
        "Copiar (los originales permanecen en el PST)",
        "Mover  (los originales se eliminan del PST)",
    };

    while (true) {
        ui.clearScreen();
        ui.printSectionTitle("Accion");
        std.debug.print("  \x1b[90mArchivo:\x1b[0m {s}\n", .{pst_path});
        std.debug.print("  \x1b[90m↑/↓ mover | Enter confirmar\x1b[0m\n\n", .{});

        for (labels, 0..) |label, idx| {
            const is_current = idx == cursor;
            if (is_current) std.debug.print("  \x1b[7m", .{});
            std.debug.print("  {s}\n", .{label});
            if (is_current) std.debug.print("  \x1b[0m", .{});
        }

        const key = ui.readSingleKey() catch continue;
        switch (key) {
            'w', 'W', 'k', 'K' => {
                if (cursor > 0) cursor -= 1;
            },
            's', 'S', 'j', 'J' => {
                if (cursor + 1 < labels.len) cursor += 1;
            },
            '1' => return "Copy",
            '2' => return "Move",
            '\r', '\n' => return if (cursor == 0) "Copy" else "Move",
            27 => {
                const seq1 = ui.readSingleKey() catch continue;
                if (seq1 != '[' and seq1 != 'O') continue;

                const seq2 = ui.readSingleKey() catch continue;
                switch (seq2) {
                    'A' => {
                        if (cursor > 0) cursor -= 1;
                    },
                    'B' => {
                        if (cursor + 1 < labels.len) cursor += 1;
                    },
                    else => {},
                }
            },
            0, 224 => {
                const ext = ui.readSingleKey() catch continue;
                switch (ext) {
                    72 => {
                        if (cursor > 0) cursor -= 1;
                    },
                    80 => {
                        if (cursor + 1 < labels.len) cursor += 1;
                    },
                    else => {},
                }
            },
            else => {},
        }
    }
}

fn chooseScanYearFilter(allocator: std.mem.Allocator) !?[]u8 {
    var cursor: usize = 0;
    const labels = [_][]const u8{
        "Todos los anios (recomendado)",
        "Un anio especifico",
    };

    while (true) {
        ui.clearScreen();
        ui.printSectionTitle("Escanear PST");
        std.debug.print("  \x1b[90mAlcance de anios para el escaneo del PST.\x1b[0m\n", .{});
        std.debug.print("  \x1b[90m↑/↓ mover | Enter confirmar | Q cancelar\x1b[0m\n\n", .{});

        for (labels, 0..) |label, idx| {
            const is_current = idx == cursor;
            if (is_current) std.debug.print("  \x1b[7m", .{});
            std.debug.print("  {s}\n", .{label});
            if (is_current) std.debug.print("  \x1b[0m", .{});
        }

        const key = ui.readSingleKey() catch continue;
        switch (key) {
            'q', 'Q' => return error.Cancelled,
            'w', 'W', 'k', 'K' => {
                if (cursor > 0) cursor -= 1;
            },
            's', 'S', 'j', 'J' => {
                if (cursor + 1 < labels.len) cursor += 1;
            },
            '\r', '\n' => {
                if (cursor == 0) return null;

                while (true) {
                    ui.clearScreen();
                    ui.printSectionTitle("Escanear PST");
                    std.debug.print("  \x1b[90mEscribe el anio para filtrar el escaneo (ej: 2023).\x1b[0m\n", .{});
                    std.debug.print("  \x1b[90mEnter vacio para volver a \"Todos los anios\".\x1b[0m\n\n", .{});
                    std.debug.print("  \x1b[33mAnio:\x1b[0m ", .{});

                    const input = ui.readLine(allocator) catch continue;
                    if (input.len == 0) {
                        allocator.free(input);
                        return null;
                    }

                    _ = std.fmt.parseInt(i32, input, 10) catch {
                        ui.printError("Anio invalido. Debe ser numerico.");
                        allocator.free(input);
                        ui.waitForEnter();
                        continue;
                    };
                    return input;
                }
            },
            27 => {
                const seq1 = ui.readSingleKey() catch continue;
                if (seq1 != '[' and seq1 != 'O') continue;

                const seq2 = ui.readSingleKey() catch continue;
                switch (seq2) {
                    'A' => {
                        if (cursor > 0) cursor -= 1;
                    },
                    'B' => {
                        if (cursor + 1 < labels.len) cursor += 1;
                    },
                    else => {},
                }
            },
            0, 224 => {
                const ext = ui.readSingleKey() catch continue;
                switch (ext) {
                    72 => {
                        if (cursor > 0) cursor -= 1;
                    },
                    80 => {
                        if (cursor + 1 < labels.len) cursor += 1;
                    },
                    else => {},
                }
            },
            else => {},
        }
    }
}

fn runScanAndSelectFolders(allocator: std.mem.Allocator, pst_path: []const u8, scan_mode: ScanMode, scan_filter_year: ?[]const u8, profile_name: ?[]const u8) !FolderSelectionResult {
    ui.clearScreen();
    ui.printSectionTitle("Escanear PST");
    std.debug.print("  \x1b[90mEjecutando escaneo {s}...\x1b[0m\n", .{if (scan_mode == .deep) "profundo" else "rapido"});

    const scan_script_path = try ps_runner.writeEmbeddedScript(allocator, .scan_pst);
    defer ps_runner.cleanupScript(allocator, scan_script_path);

    const exported_scan_path = try makeTempFilePath(allocator, "oo-scan-export", "json");
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

    var folders = std.ArrayListUnmanaged(ScannedFolder){};
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

        const item_count = extractNumber(line, "itemCount") orelse 0;
        if (item_count <= 0) continue;

        const path = extractString(line, "path") orelse continue;
        const path_copy = try allocator.dupe(u8, path);
        errdefer allocator.free(path_copy);

        const size_bytes = extractNumber(line, "sizeBytes");
        const year_breakdown_display = try extractYearBreakdownDisplay(allocator, line);
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

    const selected_flags = try promptFolderSelection(allocator, folders.items, scan_mode);
    defer allocator.free(selected_flags);

    const selected_count = countSelectedFlags(selected_flags);
    if (selected_count == 0) {
        return error.Cancelled;
    }

    const plan_path = try writeFolderPlanFromFlags(allocator, folders.items, selected_flags);

    return .{
        .folder_plan_path = plan_path,
        .selected_count = selected_count,
        .total_count = folders.items.len,
        .scan_mode = scan_mode,
    };
}

fn promptFolderSelection(allocator: std.mem.Allocator, folders: []const ScannedFolder, scan_mode: ScanMode) ![]bool {
    var tree = try buildFolderTree(allocator, folders);
    defer deinitFolderTree(allocator, &tree);

    var selected = try allocator.alloc(bool, folders.len);
    @memset(selected, false);

    var cursor: usize = 0;
    var scroll: usize = 0;
    const page_size: usize = 22;

    while (true) {
        var visible = std.ArrayListUnmanaged(VisibleTreeRow){};
        defer visible.deinit(allocator);
        try buildVisibleRows(allocator, &tree, &visible);

        if (visible.items.len == 0) return selected;
        if (cursor >= visible.items.len) cursor = visible.items.len - 1;

        if (cursor < scroll) scroll = cursor;
        if (cursor >= scroll + page_size) scroll = cursor - page_size + 1;

        ui.clearScreen();
        ui.printSectionTitle("Seleccionar carpetas");
        std.debug.print("  \x1b[90mEspacio: marcar | A: todas | C: contraer | E: expandir | W/S o ↑↓: mover | ←: contraer/desmarcar | →: expandir/marcar | Enter: confirmar | Q: cancelar\x1b[0m\n", .{});
        std.debug.print("  \x1b[90mModo: {s} | Seleccionadas: {d}/{d}\x1b[0m\n\n", .{ if (scan_mode == .deep) "Profundo" else "Rapido", countSelectedFlags(selected), folders.len });

        const start = scroll;
        const end = @min(visible.items.len, scroll + page_size);
        for (visible.items[start..end], start..) |row, row_index| {
            const node = tree.nodes.items[row.node_index];
            const is_current = row_index == cursor;
            const is_selected = if (node.folder_index) |fi| selected[fi] else false;
            const has_children = node.children.items.len > 0;
            const branch = if (has_children) (if (node.expanded) "[-]" else "[+]") else "   ";
            const mark = if (is_selected) "x" else " ";

            if (is_current) std.debug.print("  \x1b[7m", .{});
            std.debug.print("  ", .{});
            var d: usize = 0;
            while (d < row.depth) : (d += 1) {
                std.debug.print("  ", .{});
            }
            std.debug.print("{s} [{s}] {s}", .{ branch, mark, node.name });

            if (node.folder_index) |fi| {
                const folder = folders[fi];
                if (scan_mode == .deep and folder.size_bytes != null) {
                    const size_mb = @as(f64, @floatFromInt(folder.size_bytes.?)) / (1024.0 * 1024.0);
                    if (folder.year_breakdown_display.len > 0) {
                        std.debug.print(" \x1b[90m(items:{d}, anios:{s}, {d:.1}MB)\x1b[0m", .{ folder.item_count, folder.year_breakdown_display, size_mb });
                    } else {
                        std.debug.print(" \x1b[90m(items:{d}, {d:.1}MB)\x1b[0m", .{ folder.item_count, size_mb });
                    }
                } else {
                    if (folder.year_breakdown_display.len > 0) {
                        std.debug.print(" \x1b[90m(items:{d}, anios:{s})\x1b[0m", .{ folder.item_count, folder.year_breakdown_display });
                    } else {
                        std.debug.print(" \x1b[90m(items:{d})\x1b[0m", .{folder.item_count});
                    }
                }
            }
            if (is_current) std.debug.print("\x1b[0m", .{});
            std.debug.print("\n", .{});
        }

        if (end < visible.items.len) {
            std.debug.print("\n  \x1b[90m... {d} carpetas mas\x1b[0m\n", .{visible.items.len - end});
        }

        const key = ui.readSingleKey() catch continue;
        switch (key) {
            'q', 'Q' => {
                allocator.free(selected);
                return error.Cancelled;
            },
            'a', 'A' => {
                @memset(selected, true);
            },
            'c', 'C' => collapseAllNodes(&tree),
            'e', 'E' => expandAllNodes(&tree),
            'w', 'W', 'k', 'K' => {
                if (cursor > 0) cursor -= 1;
            },
            's', 'S', 'j', 'J' => {
                if (cursor + 1 < visible.items.len) cursor += 1;
            },
            ' ' => {
                const node = tree.nodes.items[visible.items[cursor].node_index];
                if (node.folder_index) |fi| {
                    selected[fi] = !selected[fi];
                }
            },
            '\r', '\n' => {
                if (countSelectedFlags(selected) == 0) continue;
                return selected;
            },
            27 => {
                const seq1 = ui.readSingleKey() catch continue;
                if (seq1 != '[' and seq1 != 'O') continue;

                const seq2 = ui.readSingleKey() catch continue;
                switch (seq2) {
                    'A' => {
                        if (cursor > 0) cursor -= 1;
                    },
                    'B' => {
                        if (cursor + 1 < visible.items.len) cursor += 1;
                    },
                    'D' => {
                        const node_index = visible.items[cursor].node_index;
                        if (tree.nodes.items[node_index].children.items.len > 0) {
                            tree.nodes.items[node_index].expanded = false;
                        } else if (tree.nodes.items[node_index].folder_index) |fi| {
                            selected[fi] = false;
                        }
                    },
                    'C' => {
                        const node_index = visible.items[cursor].node_index;
                        if (tree.nodes.items[node_index].children.items.len > 0) {
                            tree.nodes.items[node_index].expanded = true;
                        } else if (tree.nodes.items[node_index].folder_index) |fi| {
                            selected[fi] = true;
                        }
                    },
                    else => {},
                }
            },
            0, 224 => {
                const ext = ui.readSingleKey() catch continue;
                switch (ext) {
                    72 => {
                        if (cursor > 0) cursor -= 1;
                    },
                    80 => {
                        if (cursor + 1 < visible.items.len) cursor += 1;
                    },
                    75 => {
                        const node_index = visible.items[cursor].node_index;
                        if (tree.nodes.items[node_index].children.items.len > 0) {
                            tree.nodes.items[node_index].expanded = false;
                        } else if (tree.nodes.items[node_index].folder_index) |fi| {
                            selected[fi] = false;
                        }
                    },
                    77 => {
                        const node_index = visible.items[cursor].node_index;
                        if (tree.nodes.items[node_index].children.items.len > 0) {
                            tree.nodes.items[node_index].expanded = true;
                        } else if (tree.nodes.items[node_index].folder_index) |fi| {
                            selected[fi] = true;
                        }
                    },
                    else => {},
                }
            },
            else => {},
        }
    }
}

fn buildFolderTree(allocator: std.mem.Allocator, folders: []const ScannedFolder) !FolderTree {
    var tree = FolderTree{
        .nodes = .{},
        .roots = .{},
    };
    errdefer deinitFolderTree(allocator, &tree);

    for (folders, 0..) |folder, folder_index| {
        var parent: ?usize = null;
        var iter = std.mem.splitScalar(u8, folder.path, '\\');
        var prefix = std.ArrayListUnmanaged(u8){};
        defer prefix.deinit(allocator);

        while (iter.next()) |part| {
            const name = std.mem.trim(u8, part, &[_]u8{' '});
            if (name.len == 0) continue;

            if (prefix.items.len > 0) try prefix.append(allocator, '\\');
            try prefix.appendSlice(allocator, name);

            var found = findChildByName(&tree, parent, name);
            if (found == null) {
                const node_name = try allocator.dupe(u8, name);
                errdefer allocator.free(node_name);
                const node_path = try allocator.dupe(u8, prefix.items);
                errdefer allocator.free(node_path);

                const new_index = tree.nodes.items.len;
                try tree.nodes.append(allocator, .{
                    .name = node_name,
                    .full_path = node_path,
                    .parent = parent,
                    .children = .{},
                    .expanded = true,
                    .folder_index = null,
                });

                if (parent) |p| {
                    try tree.nodes.items[p].children.append(allocator, new_index);
                } else {
                    try tree.roots.append(allocator, new_index);
                }
                found = new_index;
            }

            parent = found;
        }

        if (parent) |idx| {
            tree.nodes.items[idx].folder_index = folder_index;
        }
    }

    return tree;
}

fn deinitFolderTree(allocator: std.mem.Allocator, tree: *FolderTree) void {
    for (tree.nodes.items) |*node| {
        allocator.free(node.name);
        allocator.free(node.full_path);
        node.children.deinit(allocator);
    }
    tree.nodes.deinit(allocator);
    tree.roots.deinit(allocator);
}

fn findChildByName(tree: *const FolderTree, parent: ?usize, name: []const u8) ?usize {
    if (parent) |p| {
        for (tree.nodes.items[p].children.items) |child_idx| {
            if (std.mem.eql(u8, tree.nodes.items[child_idx].name, name)) return child_idx;
        }
        return null;
    }

    for (tree.roots.items) |root_idx| {
        if (std.mem.eql(u8, tree.nodes.items[root_idx].name, name)) return root_idx;
    }
    return null;
}

fn buildVisibleRows(allocator: std.mem.Allocator, tree: *const FolderTree, out: *std.ArrayListUnmanaged(VisibleTreeRow)) !void {
    for (tree.roots.items) |root_idx| {
        try appendVisibleRowsRecursive(allocator, tree, root_idx, 0, out);
    }
}

fn appendVisibleRowsRecursive(
    allocator: std.mem.Allocator,
    tree: *const FolderTree,
    node_index: usize,
    depth: usize,
    out: *std.ArrayListUnmanaged(VisibleTreeRow),
) !void {
    try out.append(allocator, .{ .node_index = node_index, .depth = depth });
    const node = tree.nodes.items[node_index];
    if (!node.expanded) return;

    for (node.children.items) |child_idx| {
        try appendVisibleRowsRecursive(allocator, tree, child_idx, depth + 1, out);
    }
}

fn collapseAllNodes(tree: *FolderTree) void {
    for (tree.nodes.items) |*node| {
        if (node.children.items.len > 0) node.expanded = false;
    }
}

fn expandAllNodes(tree: *FolderTree) void {
    for (tree.nodes.items) |*node| {
        if (node.children.items.len > 0) node.expanded = true;
    }
}

fn countSelectedFlags(flags: []const bool) usize {
    var count: usize = 0;
    for (flags) |flag| {
        if (flag) count += 1;
    }
    return count;
}

fn writeFolderPlanFromFlags(allocator: std.mem.Allocator, folders: []const ScannedFolder, selected_flags: []const bool) ![]u8 {
    const plan_path = try makeTempFilePath(allocator, "oo-folder-plan", "json");
    errdefer allocator.free(plan_path);

    var json = std.ArrayListUnmanaged(u8){};
    defer json.deinit(allocator);

    try json.appendSlice(allocator, "{\"type\":\"folderExport\",\"folders\":[");

    var first = true;
    for (folders, 0..) |folder, idx| {
        if (!selected_flags[idx]) continue;
        if (!first) try json.appendSlice(allocator, ",");
        first = false;

        try json.appendSlice(allocator, "{\"path\":\"");
        try appendJsonEscaped(allocator, &json, folder.path);
        try json.appendSlice(allocator, "\",\"itemCount\":");
        const num_str = try std.fmt.allocPrint(allocator, "{d}", .{folder.item_count});
        defer allocator.free(num_str);
        try json.appendSlice(allocator, num_str);
        try json.appendSlice(allocator, "}");
    }

    try json.appendSlice(allocator, "]}");

    const file = std.fs.createFileAbsolute(plan_path, .{ .truncate = true }) catch {
        allocator.free(plan_path);
        return error.CannotWriteFolderPlan;
    };
    defer file.close();

    try file.writeAll(json.items);
    return plan_path;
}

fn appendJsonEscaped(allocator: std.mem.Allocator, list: *std.ArrayListUnmanaged(u8), value: []const u8) !void {
    for (value) |ch| {
        switch (ch) {
            '"' => try list.appendSlice(allocator, "\\\""),
            '\\' => try list.appendSlice(allocator, "\\\\"),
            '\n' => try list.appendSlice(allocator, "\\n"),
            '\r' => try list.appendSlice(allocator, "\\r"),
            '\t' => try list.appendSlice(allocator, "\\t"),
            else => try list.append(allocator, ch),
        }
    }
}

fn makeTempFilePath(allocator: std.mem.Allocator, prefix: []const u8, ext: []const u8) ![]u8 {
    const temp_dir = std.process.getEnvVarOwned(allocator, "TEMP") catch |err| switch (err) {
        error.EnvironmentVariableNotFound => try allocator.dupe(u8, "."),
        else => return err,
    };
    defer allocator.free(temp_dir);

    const pid: u32 = std.os.windows.GetCurrentProcessId();
    const ts: u64 = @intCast(@max(std.time.milliTimestamp(), 0));
    const file_name = try std.fmt.allocPrint(allocator, "{s}-{d}-{d}.{s}", .{ prefix, pid, ts, ext });
    defer allocator.free(file_name);

    return try std.fs.path.join(allocator, &.{ temp_dir, file_name });
}

fn cleanupTempFile(path: []const u8) void {
    if (std.fs.path.isAbsolute(path)) {
        std.fs.deleteFileAbsolute(path) catch {};
    } else {
        std.fs.cwd().deleteFile(path) catch {};
    }
}

fn executeImport(allocator: std.mem.Allocator, config: ImportConfig) !void {
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
    std.debug.print("  \x1b[1;37mTipo de buzon:\x1b[0m     {s}\n", .{store_type_friendly});
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
    var progress_state = ImportProgressState{
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
            if (extractNumber(line, "copied")) |c| last_copied = c;
            if (extractNumber(line, "moved")) |m| last_moved = m;
            if (extractNumber(line, "skipped")) |s| last_skipped = s;
            if (extractNumber(line, "failed")) |f| last_failed = f;
        } else if (std.mem.indexOf(u8, line, "\"error\"") != null) {
            error_message = extractString(line, "message") orelse line;
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
        const copied = extractNumber(json, "copied") orelse last_copied;
        const moved = extractNumber(json, "moved") orelse last_moved;
        const skipped = extractNumber(json, "skipped") orelse last_skipped;
        const failed = extractNumber(json, "failed") orelse last_failed;
        const throttle_events = extractNumber(json, "throttleEvents") orelse 0;

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

fn extractNumber(json: []const u8, key: []const u8) ?i64 {
    var search_buf: [128]u8 = undefined;
    const search = std.fmt.bufPrint(&search_buf, "\"{s}\":", .{key}) catch return null;
    const key_pos = std.mem.indexOf(u8, json, search) orelse return null;
    var pos = key_pos + search.len;

    // Skip whitespace (just in case)
    while (pos < json.len and json[pos] == ' ') : (pos += 1) {}
    if (pos >= json.len) return null;

    // Read number (may have minus sign)
    var end = pos;
    if (end < json.len and json[end] == '-') end += 1;
    while (end < json.len and json[end] >= '0' and json[end] <= '9') : (end += 1) {}
    if (end == pos) return null;

    return std.fmt.parseInt(i64, json[pos..end], 10) catch null;
}

fn extractYearBreakdownDisplay(allocator: std.mem.Allocator, json: []const u8) ![]u8 {
    const token = "\"yearBreakdown\":[";
    const start = std.mem.indexOf(u8, json, token) orelse return allocator.dupe(u8, "");

    var pos = start + token.len;
    var out = std.ArrayListUnmanaged(u8){};
    errdefer out.deinit(allocator);

    var wrote_any = false;
    while (pos < json.len) {
        if (json[pos] == ']') break;

        if (json[pos] == '{') {
            const obj_end_rel = std.mem.indexOfScalarPos(u8, json, pos, '}') orelse break;
            const row = json[pos .. obj_end_rel + 1];

            const year = extractNumber(row, "year");
            const count = extractNumber(row, "count");
            if (year != null and count != null) {
                if (wrote_any) try out.appendSlice(allocator, ", ");

                const part = try std.fmt.allocPrint(allocator, "{d}:{d}", .{ year.?, count.? });
                defer allocator.free(part);
                try out.appendSlice(allocator, part);
                wrote_any = true;
            }

            pos = obj_end_rel + 1;
            continue;
        }

        pos += 1;
    }

    if (!wrote_any) {
        out.deinit(allocator);
        return allocator.dupe(u8, "");
    }

    return try out.toOwnedSlice(allocator);
}

fn extractString(json: []const u8, key: []const u8) ?[]const u8 {
    var search_buf: [128]u8 = undefined;
    const search = std.fmt.bufPrint(&search_buf, "\"{s}\":\"", .{key}) catch return null;
    const start_pos = std.mem.indexOf(u8, json, search) orelse return null;
    const pos = start_pos + search.len;

    var end = pos;
    while (end < json.len) : (end += 1) {
        if (json[end] == '"' and (end == pos or json[end - 1] != '\\')) {
            return json[pos..end];
        }
    }
    return null;
}
