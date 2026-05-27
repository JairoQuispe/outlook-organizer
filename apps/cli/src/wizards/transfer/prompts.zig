const std = @import("std");
const ui = @import("../../ui.zig");
const types = @import("types.zig");

pub fn chooseImportAction() []const u8 {
    var cursor: usize = 0;
    const labels = [_][]const u8{
        "Copiar (los originales permanecen en el origen)",
        "Mover (los originales se eliminan del origen)",
    };

    while (true) {
        ui.clearScreen();
        ui.printSectionTitle("Accion");
        std.debug.print("  \x1b[90mW/S o ↑/↓ mover | Enter confirmar | 1/2 seleccionar\x1b[0m\n\n", .{});

        for (labels, 0..) |label, idx| {
            const is_current = idx == cursor;
            if (is_current) std.debug.print("  \x1b[7m", .{});
            std.debug.print("  {s}", .{label});
            if (is_current) std.debug.print("\x1b[0m", .{});
            std.debug.print("\n", .{});
        }

        const input = ui.readMenuInput(&cursor, labels.len) catch continue;
        switch (input) {
            .cancel => return "Copy",
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
