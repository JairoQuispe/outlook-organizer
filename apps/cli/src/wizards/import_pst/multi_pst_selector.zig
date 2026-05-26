const std = @import("std");
const ui = @import("../../ui.zig");
const file_browser = @import("../../file_browser.zig");

/// Estructura para almacenar la lista de PSTs seleccionados
pub const PstListResult = struct {
    paths: [][]u8,
    allocator: std.mem.Allocator,

    pub fn deinit(self: PstListResult) void {
        for (self.paths) |path| {
            self.allocator.free(path);
        }
        self.allocator.free(self.paths);
    }
};

fn containsPath(paths: [][]u8, target: []const u8) bool {
    for (paths) |existing| {
        if (std.mem.eql(u8, existing, target)) return true;
    }
    return false;
}

/// Permite seleccionar múltiples archivos PST de forma interactiva
pub fn selectMultiplePstFiles(allocator: std.mem.Allocator) !PstListResult {
    var paths = std.ArrayListUnmanaged([]u8){};
    errdefer {
        for (paths.items) |p| allocator.free(p);
        paths.deinit(allocator);
    }

    while (true) {
        ui.clearScreen();
        ui.printSectionTitle("Seleccion de archivos PST");
        std.debug.print("  \x1b[1;32mPSTs seleccionados actualmente ({d}):\x1b[0m\n", .{paths.items.len});
        if (paths.items.len == 0) {
            std.debug.print("    \x1b[90mNinguno. Debes seleccionar al menos uno.\x1b[0m\n", .{});
        } else {
            for (paths.items) |p| {
                std.debug.print("    - \x1b[32m{s}\x1b[0m\n", .{p});
            }
        }
        std.debug.print("\n  \x1b[1;37mOpciones:\x1b[0m\n", .{});
        std.debug.print("    \x1b[33m[1]\x1b[0m Agregar un archivo PST\n", .{});
        if (paths.items.len > 0) {
            std.debug.print("    \x1b[33m[2]\x1b[0m Quitar el ultimo PST agregado\n", .{});
            std.debug.print("    \x1b[33m[3]\x1b[0m Finalizar seleccion ({d} PSTs) y continuar\n", .{paths.items.len});
        }
        std.debug.print("    \x1b[33m[Q]\x1b[0m Cancelar operacion\n\n", .{});
        std.debug.print("  \x1b[33mSeleccione una opcion:\x1b[0m ", .{});

        const key = try ui.readSingleKey();
        switch (key) {
            'q', 'Q' => return error.Cancelled,
            '1' => {
                const p = file_browser.selectPstFile(allocator) catch |err| {
                    if (err == error.Cancelled) continue;
                    return err;
                };
                if (containsPath(paths.items, p)) {
                    allocator.free(p);
                    ui.printError("Ese archivo PST ya ha sido seleccionado");
                    ui.waitForEnter();
                } else {
                    try paths.append(allocator, p);
                }
            },
            '2' => {
                if (paths.items.len > 0) {
                    const last = paths.pop().?;
                    allocator.free(last);
                }
            },
            '3' => {
                if (paths.items.len > 0) {
                    break;
                }
            },
            else => {},
        }
    }

    return PstListResult{
        .paths = try paths.toOwnedSlice(allocator),
        .allocator = allocator,
    };
}
