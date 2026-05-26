const std = @import("std");
const ui = @import("../../ui.zig");
const store_selector = @import("../../store_selector.zig");
const types = @import("types.zig");
const utils = @import("utils.zig");
const menu_ui = @import("menu_ui.zig");

pub const TempRoutingSelection = struct {
    year: i32,
    month: ?u8,
    selected: bool,
};

const DateItem = struct {
    year: i32,
    month: u8,
};

fn dateKey(year: i32, month: u8) ?u32 {
    const year_u32 = std.math.cast(u32, year) orelse return null;
    return (year_u32 << 8) | @as(u32, month);
}

fn compareDateItemsDesc(context: void, a: DateItem, b: DateItem) bool {
    _ = context;
    if (a.year != b.year) return a.year > b.year;
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

fn renderCriterionMenu(cursor: usize, labels: []const []const u8) void {
    menu_ui.beginMenu("Criterio de Enrutamiento");
    menu_ui.printMutedLine("Selecciona como deseas agrupar los correos para enviarlos a diferentes buzones.");
    menu_ui.printMutedLine("↑/↓ mover | Enter confirmar | Q cancelar");
    std.debug.print("\n", .{});

    for (labels, 0..) |label, idx| {
        menu_ui.printSelectableLabel(label, idx == cursor);
    }
}

fn unassignedMapping(allocator: std.mem.Allocator, year: i32, month: ?u8) !types.TargetStoreMapping {
    return .{
        .year = year,
        .month = month,
        .store_id = try allocator.dupe(u8, ""),
        .store_name = try allocator.dupe(u8, "Sin asignar (se omitira)"),
        .store_type = try allocator.dupe(u8, ""),
    };
}

fn renderRoutingOptions(cursor: usize, criterion: types.RoutingCriterion, options: []const types.TargetStoreMapping) void {
    menu_ui.beginMenu("Enrutamiento de Correos");
    menu_ui.printMutedLine("Asigna un buzon de destino para cada grupo de fecha disponible en el PST.");
    menu_ui.printMutedLine("↑/↓ mover | Enter asignar buzon | D desactivar/quitar | F finalizar mapeo | Q cancelar");
    std.debug.print("\n", .{});

    for (options, 0..) |opt, idx| {
        const is_current = idx == cursor;
        menu_ui.beginHighlightedRow(is_current);

        if (criterion == .by_year) {
            std.debug.print("  Anio {d:4} => {s}", .{ opt.year, opt.store_name });
        } else {
            const month_name = getMonthName(opt.month.?);
            std.debug.print("  {s:9} {d:4} => {s}", .{ month_name, opt.year, opt.store_name });
        }

        menu_ui.endHighlightedRow(is_current);
        std.debug.print("\n", .{});
    }
}

fn assignStoreToOption(allocator: std.mem.Allocator, opt: *types.TargetStoreMapping, selected_store: anytype) void {
    allocator.free(opt.store_id);
    allocator.free(opt.store_name);
    allocator.free(opt.store_type);
    opt.store_id = selected_store.store_id;
    opt.store_name = selected_store.display_name;
    opt.store_type = selected_store.store_type;
}

fn resetOption(allocator: std.mem.Allocator, opt: *types.TargetStoreMapping) !void {
    allocator.free(opt.store_id);
    allocator.free(opt.store_name);
    allocator.free(opt.store_type);
    opt.* = try unassignedMapping(allocator, opt.year, opt.month);
}

fn buildRoutingOptions(
    allocator: std.mem.Allocator,
    criterion: types.RoutingCriterion,
    items_found: []DateItem,
) ![]types.TargetStoreMapping {
    var options = std.ArrayListUnmanaged(types.TargetStoreMapping){};
    errdefer {
        for (options.items) |opt| {
            opt.deinit(allocator);
        }
        options.deinit(allocator);
    }

    if (criterion == .by_year) {
        var seen_years = std.AutoHashMap(i32, void).init(allocator);
        defer seen_years.deinit();

        var years = std.ArrayListUnmanaged(i32){};
        defer years.deinit(allocator);

        for (items_found) |it| {
            if (seen_years.contains(it.year)) continue;
            try seen_years.put(it.year, {});
            try years.append(allocator, it.year);
        }

        std.sort.pdq(i32, years.items, {}, std.sort.desc(i32));

        for (years.items) |year| {
            try options.append(allocator, try unassignedMapping(allocator, year, null));
        }
    } else {
        std.sort.pdq(DateItem, items_found, {}, compareDateItemsDesc);
        for (items_found) |it| {
            try options.append(allocator, try unassignedMapping(allocator, it.year, it.month));
        }
    }

    return try options.toOwnedSlice(allocator);
}

pub fn selectRoutingCriterion() !types.RoutingCriterion {
    var cursor: usize = 0;
    const labels = [_][]const u8{
        "Agrupado por Anos  (Ej: 2023 -> Buzon A, 2024 -> Buzon B)",
        "Agrupado por Meses (Ej: Enero 2024 -> Buzon A, Febrero 2024 -> Buzon B)",
    };

    while (true) {
        renderCriterionMenu(cursor, labels[0..]);

        const input = ui.readMenuInput(&cursor, labels.len) catch continue;
        switch (input) {
            .cancel => return error.Cancelled,
            .enter => return if (cursor == 0) .by_year else .by_month,
            else => {},
        }
    }
}

pub fn configureMappings(
    allocator: std.mem.Allocator,
    criterion: types.RoutingCriterion,
    folders_json_path: []const u8,
    profile_name: ?[]const u8,
) ![]types.TargetStoreMapping {
    const items_found = try parseAvailableDates(allocator, folders_json_path);
    defer allocator.free(items_found);

    if (items_found.len == 0) {
        ui.printError("No se encontraron correos con fecha en el escaneo previo para enrutar.");
        ui.waitForEnter();
        return error.NoDataToRoute;
    }

    const options = try buildRoutingOptions(allocator, criterion, items_found);
    var success = false;
    defer if (!success) {
        utils.freeTargetStoreMappings(allocator, options);
    };

    var cursor: usize = 0;
    while (true) {
        renderRoutingOptions(cursor, criterion, options);

        const input = ui.readMenuInput(&cursor, options.len) catch continue;
        switch (input) {
            .cancel => return error.Cancelled,
            .enter => {
                const selected_store = store_selector.selectTargetStore(allocator, profile_name) catch |err| {
                    if (err == error.Cancelled) continue;
                    ui.printError("Error seleccionando buzon");
                    ui.waitForEnter();
                    continue;
                };
                assignStoreToOption(allocator, &options[cursor], selected_store);
            },
            .key => |key| switch (key) {
                'f', 'F' => {
                    if (utils.countAssignedMappings(options) == 0) {
                        ui.printError("Debes asignar al menos un buzon de destino para continuar.");
                        ui.waitForEnter();
                        continue;
                    }
                    success = true;
                    return options;
                },
                'd', 'D' => try resetOption(allocator, &options[cursor]),
                else => {},
            },
            else => {},
        }
    }
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

    var seen_dates = std.AutoHashMap(u32, void).init(allocator);
    defer seen_dates.deinit();

    var month_pos: usize = 0;
    while (std.mem.indexOfPos(u8, buffer, month_pos, "\"month\"")) |key_pos| {
        month_pos = key_pos + "\"month\"".len;
        const row = utils.findEnclosingObject(buffer, key_pos) orelse continue;
        const month_str = utils.extractString(row, "month") orelse continue;
        const count = utils.extractNumber(row, "count") orelse 0;
        if (count <= 0) continue;
        if (month_str.len < 7 or month_str[4] != '-') continue;

        const year = std.fmt.parseInt(i32, month_str[0..4], 10) catch continue;
        const month = std.fmt.parseInt(u8, month_str[5..7], 10) catch continue;
        if (month < 1 or month > 12) continue;

        const key = dateKey(year, month) orelse continue;
        if (seen_dates.contains(key)) continue;
        try seen_dates.put(key, {});
        try items.append(allocator, .{ .year = year, .month = month });
    }

    if (items.items.len > 0) {
        return try items.toOwnedSlice(allocator);
    }

    var year_pos: usize = 0;
    while (std.mem.indexOfPos(u8, buffer, year_pos, "\"year\"")) |key_pos| {
        year_pos = key_pos + "\"year\"".len;
        const row = utils.findEnclosingObject(buffer, key_pos) orelse continue;
        const year_num = utils.extractNumber(row, "year") orelse continue;
        const count = utils.extractNumber(row, "count") orelse 0;
        if (count <= 0) continue;

        const year = std.math.cast(i32, year_num) orelse continue;
        if (year < 1900 or year > 9999) continue;

        var m: u8 = 1;
        while (m <= 12) : (m += 1) {
            const key = dateKey(year, m) orelse continue;
            if (seen_dates.contains(key)) continue;
            try seen_dates.put(key, {});
            try items.append(allocator, .{ .year = year, .month = m });
        }
    }

    return try items.toOwnedSlice(allocator);
}
