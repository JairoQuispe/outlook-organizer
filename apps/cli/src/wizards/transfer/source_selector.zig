const std = @import("std");
const ui = @import("../../ui.zig");
const store_selector = @import("../../store_selector.zig");
const file_browser = @import("../../file_browser.zig");
const types = @import("types.zig");

pub fn selectSource(allocator: std.mem.Allocator, profile_name: ?[]const u8) !types.SourceInfo {
    var cursor: usize = 0;
    const labels = [_][]const u8{
        "Mi buzon de Exchange (predeterminado)",
        "Otro buzon Exchange Online / OST",
        "Archivo PST",
    };

    while (true) {
        ui.clearScreen();
        ui.printSectionTitle("Origen de los correos");
        std.debug.print("  \x1b[90mW/S o ↑/↓ mover | Enter confirmar | Q cancelar\x1b[0m\n\n", .{});

        for (labels, 0..) |label, idx| {
            const is_current = idx == cursor;
            if (is_current) std.debug.print("  \x1b[7m", .{});
            std.debug.print("  {s}", .{label});
            if (is_current) std.debug.print("\x1b[0m", .{});
            std.debug.print("\n", .{});
        }

        const input = ui.readMenuInput(&cursor, labels.len) catch continue;
        switch (input) {
            .cancel => return error.Cancelled,
            .enter => {
                if (cursor == 0) {
                    return types.SourceInfo{
                        .store_id = try allocator.dupe(u8, ""),
                        .store_name = try allocator.dupe(u8, "Buzon predeterminado"),
                        .store_type = try allocator.dupe(u8, "ExchangeOnline"),
                        .pst_path = try allocator.dupe(u8, ""),
                    };
                } else if (cursor == 1) {
                    const chosen = try store_selector.selectTargetStore(allocator, profile_name);
                    return types.SourceInfo{
                        .store_id = chosen.store_id,
                        .store_name = chosen.display_name,
                        .store_type = chosen.store_type,
                        .pst_path = try allocator.dupe(u8, ""),
                    };
                } else {
                    const pst = try file_browser.selectPstFile(allocator);
                    return types.SourceInfo{
                        .store_id = try allocator.dupe(u8, ""),
                        .store_name = try allocator.dupe(u8, pst),
                        .store_type = try allocator.dupe(u8, "PST"),
                        .pst_path = pst,
                    };
                }
            },
            else => {},
        }
    }
}
