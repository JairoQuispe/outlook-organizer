const std = @import("std");
const ui = @import("../../ui.zig");
const file_browser = @import("../../file_browser.zig");
const multi_pst_selector = @import("multi_pst_selector.zig");
const single_import_flow = @import("flow/single_import_flow.zig");
const batch_import_flow = @import("flow/batch_import_flow.zig");

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
                ui.cancelAbort();
                return;
            }
            ui.failAbort("Error seleccionando archivos PST");
            return;
        };
    } else {
        single_pst_path = file_browser.selectPstFile(allocator) catch |err| {
            if (err == error.Cancelled) {
                ui.cancelAbort();
                return;
            }
            ui.failAbort("Error seleccionando archivo PST");
            return;
        };
    }

    const pst_paths: [][]const u8 = if (is_multi_pst) blk: {
        const slice = try allocator.alloc([]const u8, psts_result.?.paths.len);
        for (psts_result.?.paths, 0..) |path, i| {
            slice[i] = path;
        }
        break :blk slice;
    } else blk: {
        const slice = try allocator.alloc([]const u8, 1);
        slice[0] = single_pst_path.?;
        break :blk slice;
    };
    defer allocator.free(pst_paths);

    if (is_multi_pst) {
        try batch_import_flow.runMultiPstBatch(allocator, pst_paths);
        return;
    }

    try single_import_flow.runSingleImport(allocator, pst_paths[0]);
}
