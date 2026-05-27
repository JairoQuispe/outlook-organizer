const std = @import("std");
const ui = @import("../../ui.zig");
const single_flow = @import("flow/single_transfer_flow.zig");
const batch_flow = @import("flow/batch_transfer_flow.zig");

pub fn run(allocator: std.mem.Allocator) !void {
    ui.clearScreen();
    ui.printSectionTitle("Tipo de Transferencia");
    std.debug.print("  \x1b[33m[1]\x1b[0m Un unico origen hacia uno o multiples destinos\n", .{});
    std.debug.print("  \x1b[33m[2]\x1b[0m Multiples origenes hacia uno o multiples destinos (Lote secuencial)\n\n", .{});
    std.debug.print("  \x1b[33mSeleccione el modo de operacion [1]:\x1b[0m ", .{});
    const mode_key = ui.readSingleKey() catch '1';
    const is_multi = (mode_key == '2');

    ui.clearScreen();

    if (is_multi) {
        try batch_flow.runMultiSourceBatch(allocator);
    } else {
        try single_flow.runSingleTransfer(allocator);
    }
}
