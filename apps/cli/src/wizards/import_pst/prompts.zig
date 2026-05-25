const std = @import("std");
const ui = @import("../../ui.zig");
const types = @import("types.zig");

pub fn chooseOutlookProfile(allocator: std.mem.Allocator) !?[]const u8 {
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

pub fn chooseScanMode(allocator: std.mem.Allocator) !types.ScanMode {
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

pub fn chooseImportAction(pst_path: []const u8) []const u8 {
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

pub fn chooseScanYearFilter(allocator: std.mem.Allocator) !?[]u8 {
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
