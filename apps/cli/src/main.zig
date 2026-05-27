const std = @import("std");
const ui = @import("ui.zig");
const import_wizard = @import("wizards/import_pst.zig");
const transfer_wizard = @import("wizards/transfer.zig");

pub fn main() !void {
    var gpa: std.heap.GeneralPurposeAllocator(.{}) = .init;
    defer _ = gpa.deinit();
    const allocator = gpa.allocator();

    ui.ensureAnsiMode();

    while (true) {
        const choice = ui.readMenuChoice() catch {
            continue;
        };

        switch (choice) {
            .import_pst => {
                import_wizard.run(allocator) catch {
                    ui.printError("Error en el asistente de importacion");
                    ui.waitForEnter();
                };
            },
            .scan_pst => {
                ui.clearScreen();
                std.debug.print("\n  Scan PST - Proximamente\n\n", .{});
                ui.waitForEnter();
            },
            .transfer => {
                transfer_wizard.run(allocator) catch |err| {
                    if (err == error.Cancelled) {
                        ui.cancelAbort();
                    } else {
                        ui.printError("Error en el asistente de transferencia");
                        ui.waitForEnter();
                    }
                };
            },
            .exit => {
                std.debug.print("\n  Hasta luego.\n\n", .{});
                break;
            },
        }
    }
}
