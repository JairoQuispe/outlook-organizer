const std = @import("std");
const ui = @import("../../ui.zig");
const store_selector = @import("../../store_selector.zig");
const types = @import("types.zig");

pub fn selectRoutingCriterion() !types.RoutingCriterion {
    var cursor: usize = 0;
    const labels = [_][]const u8{
        "Agrupado por Anos  (Ej: 2023 -> Buzon A, 2024 -> Buzon B)",
        "Agrupado por Meses (Ej: Enero 2024 -> Buzon A, Febrero 2024 -> Buzon B)",
    };

    while (true) {
        ui.clearScreen();
        ui.printSectionTitle("Criterio de Enrutamiento");
        std.debug.print("  \x1b[90mSelecciona como deseas agrupar los correos para enviarlos a diferentes buzones.\x1b[0m\n", .{});
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
                return if (cursor == 0) .by_year else .by_month;
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

pub const TempRoutingSelection = struct {
    year: i32,
    month: ?u8,
    selected: bool,
};

pub fn configureMappings(
    allocator: std.mem.Allocator,
    criterion: types.RoutingCriterion,
    folders_json_path: []const u8,
    profile_name: ?[]const u8,
) ![]types.TargetStoreMapping {
    // 1. Leer el archivo JSON para saber qué años y meses hay disponibles
    const items_found = try parseAvailableDates(allocator, folders_json_path);
    defer {
        for (items_found) |item| {
            _ = item;
        }
        allocator.free(items_found);
    }

    if (items_found.len == 0) {
        ui.printError("No se encontraron correos con fecha en el escaneo previo para enrutar.");
        ui.waitForEnter();
        return error.NoDataToRoute;
    }

    // 2. Generar la lista de opciones según el criterio
    var options = std.ArrayListUnmanaged(types.TargetStoreMapping){};
    errdefer {
        for (options.items) |opt| {
            allocator.free(opt.store_id);
            allocator.free(opt.store_name);
            allocator.free(opt.store_type);
        }
        options.deinit(allocator);
    }

    if (criterion == .by_year) {
        // Encontrar años únicos
        var years = std.ArrayListUnmanaged(i32){};
        defer years.deinit(allocator);
        for (items_found) |it| {
            var exists = false;
            for (years.items) |y| {
                if (y == it.year) {
                    exists = true;
                    break;
                }
            }
            if (!exists) {
                try years.append(allocator, it.year);
            }
        }
        // Ordenar años de forma descendente
        std.sort.pdq(i32, years.items, {}, std.sort.desc(i32));

        for (years.items) |year| {
            try options.append(allocator, .{
                .year = year,
                .month = null,
                .store_id = try allocator.dupe(u8, ""),
                .store_name = try allocator.dupe(u8, "Sin asignar (se omitira)"),
                .store_type = try allocator.dupe(u8, ""),
            });
        }
    } else {
        // Ordenar items por año desc, mes desc
        std.sort.pdq(DateItem, items_found, {}, compareDateItemsDesc);

        for (items_found) |it| {
            try options.append(allocator, .{
                .year = it.year,
                .month = it.month,
                .store_id = try allocator.dupe(u8, ""),
                .store_name = try allocator.dupe(u8, "Sin asignar (se omitira)"),
                .store_type = try allocator.dupe(u8, ""),
            });
        }
    }

    // 3. Menú interactivo para mapear cada opción a un buzón
    var cursor: usize = 0;
    while (true) {
        ui.clearScreen();
        ui.printSectionTitle("Enrutamiento de Correos");
        std.debug.print("  \x1b[90mAsigna un buzon de destino para cada grupo de fecha disponible en el PST.\x1b[0m\n", .{});
        std.debug.print("  \x1b[90m↑/↓ mover | Enter asignar buzon | D desactivar/quitar | F finalizar mapeo | Q cancelar\x1b[0m\n\n", .{});

        for (options.items, 0..) |opt, idx| {
            const is_current = idx == cursor;
            if (is_current) std.debug.print("  \x1b[7m", .{});

            if (criterion == .by_year) {
                std.debug.print("  Anio {d:4} => {s}", .{ opt.year, opt.store_name });
            } else {
                const month_name = getMonthName(opt.month.?);
                std.debug.print("  {s:9} {d:4} => {s}", .{ month_name, opt.year, opt.store_name });
            }

            if (is_current) std.debug.print("\x1b[0m", .{});
            std.debug.print("\n", .{});
        }

        const key = ui.readSingleKey() catch continue;
        switch (key) {
            'q', 'Q' => return error.Cancelled,
            'f', 'F' => {
                // Verificar si al menos hay un mapeo configurado
                var mapped_count: usize = 0;
                for (options.items) |opt| {
                    if (opt.store_id.len > 0) mapped_count += 1;
                }
                if (mapped_count == 0) {
                    ui.printError("Debes asignar al menos un buzon de destino para continuar.");
                    ui.waitForEnter();
                    continue;
                }
                break;
            },
            'd', 'D' => {
                // Quitar asignación
                var opt = &options.items[cursor];
                allocator.free(opt.store_id);
                allocator.free(opt.store_name);
                allocator.free(opt.store_type);
                opt.store_id = try allocator.dupe(u8, "");
                opt.store_name = try allocator.dupe(u8, "Sin asignar (se omitira)");
                opt.store_type = try allocator.dupe(u8, "");
            },
            'w', 'W', 'k', 'K' => {
                if (cursor > 0) cursor -= 1;
            },
            's', 'S', 'j', 'J' => {
                if (cursor + 1 < options.items.len) cursor += 1;
            },
            '\r', '\n' => {
                // Seleccionar buzón de destino para el elemento actual
                const selected_store = store_selector.selectTargetStore(allocator, profile_name) catch |err| {
                    if (err == error.Cancelled) continue;
                    ui.printError("Error seleccionando buzon");
                    ui.waitForEnter();
                    continue;
                };
                var opt = &options.items[cursor];
                allocator.free(opt.store_id);
                allocator.free(opt.store_name);
                allocator.free(opt.store_type);
                opt.store_id = selected_store.store_id; // Toma posesión
                opt.store_name = selected_store.display_name; // Toma posesión
                opt.store_type = selected_store.store_type; // Toma posesión
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
                        if (cursor + 1 < options.items.len) cursor += 1;
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
                        if (cursor + 1 < options.items.len) cursor += 1;
                    },
                    else => {},
                }
            },
            else => {},
        }
    }

    return try options.toOwnedSlice(allocator);
}

const DateItem = struct {
    year: i32,
    month: u8,
};

fn compareDateItemsDesc(context: void, a: DateItem, b: DateItem) bool {
    _ = context;
    if (a.year != b.year) {
        return a.year > b.year;
    }
    return a.month > b.month;
}

fn getMonthName(month: u8) []const u8 {
    return switch (month) {
        1 => "Enero",
        2 => "Febrero",
        3 => "Marzo",
        4 => "Abril",
        5 => "Mayo",
        6 => "Junio",
        7 => "Julio",
        8 => "Agosto",
        9 => "Septiembre",
        10 => "Octubre",
        11 => "Noviembre",
        12 => "Diciembre",
        else => "Desconocido",
    };
}

fn parseAvailableDates(allocator: std.mem.Allocator, json_path: []const u8) ![]DateItem {
    var file = try std.fs.openFileAbsolute(json_path, .{});
    defer file.close();

    const file_size = (try file.stat()).size;
    const buffer = try allocator.alloc(u8, file_size);
    defer allocator.free(buffer);

    _ = try file.readAll(buffer);

    var items = std.ArrayListUnmanaged(DateItem){};
    errdefer items.deinit(allocator);

    // Parsear el JSON manualmente de forma sencilla extrayendo los yearBreakdown de las carpetas
    var pos: usize = 0;
    const token = "\"yearBreakdown\":[";
    while (std.mem.indexOfPos(u8, buffer, pos, token)) |start| {
        pos = start + token.len;
        while (pos < buffer.len) {
            if (buffer[pos] == ']') break;
            if (buffer[pos] == '{') {
                const obj_end = std.mem.indexOfScalarPos(u8, buffer, pos, '}') orelse break;
                const row = buffer[pos .. obj_end + 1];

                const year = extractNumber(row, "year");
                if (year) |y| {
                    // Por simplicidad, agregamos meses para ese año (por defecto asumimos el mes si se requiere, pero podemos escanear las estadísticas o simplemente registrar el año)
                    // Para soportar meses detallados, vamos a buscar si hay un "monthBreakdown" en el JSON
                    // El scan con -ExportStatistics incluye monthBreakdown en cada carpeta. Busquemos si está presente.
                    const month_token = "\"monthBreakdown\":[";
                    if (std.mem.indexOfPos(u8, buffer, start, month_token)) |m_start| {
                        // Si hay monthBreakdown, parseamos los meses reales
                        var m_pos = m_start + month_token.len;
                        while (m_pos < buffer.len) {
                            if (buffer[m_pos] == ']') break;
                            if (buffer[m_pos] == '{') {
                                const m_obj_end = std.mem.indexOfScalarPos(u8, buffer, m_pos, '}') orelse break;
                                const m_row = buffer[m_pos .. m_obj_end + 1];
                                const m_str = extractString(m_row, "month"); // "2024-03"
                                const m_count = extractNumber(m_row, "count") orelse 0;

                                if (m_str != null and m_count > 0) {
                                    // Parsear año y mes de "2024-03"
                                    if (m_str.?.len >= 7 and m_str.?[4] == '-') {
                                        const parsed_y = std.fmt.parseInt(i32, m_str.?[0..4], 10) catch @as(i32, @intCast(y));
                                        const parsed_m = std.fmt.parseInt(u8, m_str.?[5..7], 10) catch @as(u8, 1);

                                        // Evitar duplicados en la lista temporal
                                        var exists = false;
                                        for (items.items) |it| {
                                            if (it.year == parsed_y and it.month == parsed_m) {
                                                exists = true;
                                                break;
                                            }
                                        }
                                        if (!exists) {
                                            try items.append(allocator, .{ .year = parsed_y, .month = parsed_m });
                                        }
                                    }
                                }
                                m_pos = m_obj_end + 1;
                                continue;
                            }
                            m_pos += 1;
                        }
                    } else {
                        // Si no hay mes detallado (por ej. escaneo rápido), registramos todos los meses del 1 al 12
                        // Aunque para máxima precisión, el wizard forzará un escaneo con estadísticas si se selecciona enrutamiento.
                        var m: u8 = 1;
                        while (m <= 12) : (m += 1) {
                            var exists = false;
                            for (items.items) |it| {
                                if (it.year == y and it.month == m) {
                                    exists = true;
                                    break;
                                }
                            }
                            if (!exists) {
                                try items.append(allocator, .{ .year = @intCast(y), .month = m });
                            }
                        }
                    }
                }
                pos = obj_end + 1;
                continue;
            }
            pos += 1;
        }
    }

    return try items.toOwnedSlice(allocator);
}

fn extractNumber(json: []const u8, key: []const u8) ?i64 {
    var search_buf: [128]u8 = undefined;
    const search = std.fmt.bufPrint(&search_buf, "\"{s}\":", .{key}) catch return null;
    const key_pos = std.mem.indexOf(u8, json, search) orelse return null;
    var pos = key_pos + search.len;
    while (pos < json.len and json[pos] == ' ') : (pos += 1) {}
    if (pos >= json.len) return null;
    var end = pos;
    if (end < json.len and json[end] == '-') end += 1;
    while (end < json.len and json[end] >= '0' and json[end] <= '9') : (end += 1) {}
    if (end == pos) return null;
    return std.fmt.parseInt(i64, json[pos..end], 10) catch null;
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
