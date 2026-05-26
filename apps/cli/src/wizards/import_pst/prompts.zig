const std = @import("std");
const ui = @import("../../ui.zig");
const types = @import("types.zig");
const menu_ui = @import("menu_ui.zig");

fn renderMenu(
    title: []const u8,
    context_label: ?[]const u8,
    context_value: ?[]const u8,
    intro_lines: []const []const u8,
    labels: []const []const u8,
    cursor: usize,
) void {
    menu_ui.beginMenu(title);

    if (context_label) |label| {
        menu_ui.printContext(label, context_value orelse "");
    }

    for (intro_lines) |line| {
        menu_ui.printMutedLine(line);
    }
    std.debug.print("\n", .{});

    for (labels, 0..) |label, idx| {
        menu_ui.printSelectableLabel(label, idx == cursor);
    }
}

pub fn chooseOutlookProfile(allocator: std.mem.Allocator) !?[]const u8 {
    var cursor: usize = 0;
    const labels = [_][]const u8{
        "Usar perfil predeterminado de Outlook",
        "Especificar nombre del perfil",
    };

    while (true) {
        const intro_lines = [_][]const u8{
            "Selecciona el perfil de Outlook que deseas usar para esta importacion.",
            "W/S o ↑/↓ mover | Enter confirmar",
        };
        renderMenu("Seleccionar perfil de Outlook", null, null, intro_lines[0..], labels[0..], cursor);

        const input = ui.readMenuInput(&cursor, labels.len) catch continue;
        switch (input) {
            .enter => {
                if (cursor == 0) {
                    return null;
                }

                while (true) {
                    ui.clearScreen();
                    ui.printSectionTitle("Especificar perfil de Outlook");
                    std.debug.print("  \x1b[90mEscribe el nombre del perfil de Outlook que deseas usar.\x1b[0m\n\n", .{});
                    std.debug.print("  \x1b[33mNombre del perfil:\x1b[0m ", .{});

                    const profile_input = ui.readLine(allocator) catch continue;
                    if (profile_input.len == 0) {
                        allocator.free(profile_input);
                        continue;
                    }
                    return profile_input;
                }
            },
            .key => |_| {},
            else => {},
        }
    }
}

pub fn chooseScanMode(allocator: std.mem.Allocator) !types.ScanMode {
    _ = allocator;
    var cursor: usize = 0;

    while (true) {
        const labels = [_][]const u8{
            "Escaneo Rapido  (carpetas / conteo de items)",
            "Escaneo Profundo (carpetas / conteo de items / tamano)",
        };
        const intro_lines = [_][]const u8{
            "Selecciona el tipo de escaneo antes de definir la accion.",
            "W/S o ↑/↓ mover | Enter confirmar",
        };
        renderMenu("Escanear PST", null, null, intro_lines[0..], labels[0..], cursor);

        const input = ui.readMenuInput(&cursor, labels.len) catch continue;
        switch (input) {
            .key => |key| switch (key) {
                '1' => return .quick,
                '2' => return .deep,
                else => {},
            },
            .enter => return if (cursor == 0) .quick else .deep,
            else => {},
        }
    }
}

pub fn chooseImportAction(pst_path: []const u8) []const u8 {
    var cursor: usize = 0;
    const labels = [_][]const u8{
        "Copiar (los originales permanecen en el PST)",
        "Mover  (los originales se eliminan del PST)",
    };

    while (true) {
        const intro_lines = [_][]const u8{"W/S o ↑/↓ mover | Enter confirmar"};
        renderMenu("Accion", "Archivo", pst_path, intro_lines[0..], labels[0..], cursor);

        const input = ui.readMenuInput(&cursor, labels.len) catch continue;
        switch (input) {
            .key => |key| switch (key) {
                '1' => return "Copy",
                '2' => return "Move",
                else => {},
            },
            .enter => return if (cursor == 0) "Copy" else "Move",
            else => {},
        }
    }
}

pub fn chooseScanYearFilter(allocator: std.mem.Allocator) !?[]u8 {
    var cursor: usize = 0;
    const labels = [_][]const u8{
        "Todos los anios (recomendado)",
        "Un anio especifico",
    };

    while (true) {
        const intro_lines = [_][]const u8{
            "Alcance de anios para el escaneo del PST.",
            "W/S o ↑/↓ mover | Enter confirmar | Q cancelar",
        };
        renderMenu("Escanear PST", null, null, intro_lines[0..], labels[0..], cursor);

        const input = ui.readMenuInput(&cursor, labels.len) catch continue;
        switch (input) {
            .cancel => return error.Cancelled,
            .enter => {
                if (cursor == 0) return null;

                while (true) {
                    ui.clearScreen();
                    ui.printSectionTitle("Escanear PST");
                    std.debug.print("  \x1b[90mEscribe el anio para filtrar el escaneo (ej: 2023).\x1b[0m\n", .{});
                    std.debug.print("  \x1b[90mEnter vacio para volver a \"Todos los anios\".\x1b[0m\n\n", .{});
                    std.debug.print("  \x1b[33mAnio:\x1b[0m ", .{});

                    const year_input = ui.readLine(allocator) catch continue;
                    if (year_input.len == 0) {
                        allocator.free(year_input);
                        return null;
                    }

                    _ = std.fmt.parseInt(i32, year_input, 10) catch {
                        ui.printError("Anio invalido. Debe ser numerico.");
                        allocator.free(year_input);
                        ui.waitForEnter();
                        continue;
                    };
                    return year_input;
                }
            },
            .key => |_| {},
            else => {},
        }
    }
}
