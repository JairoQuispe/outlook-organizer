const std = @import("std");
const ui = @import("../../ui.zig");
const store_selector = @import("../../store_selector.zig");
const file_browser = @import("../../file_browser.zig");
const types = @import("types.zig");

pub fn selectDestination(allocator: std.mem.Allocator, profile_name: ?[]const u8) !types.DestInfo {
    var cursor: usize = 0;
    const labels = [_][]const u8{
        "Mi buzon de Exchange (predeterminado)",
        "Otro buzon Exchange Online / OST",
        "Crear o usar archivo PST",
    };

    while (true) {
        ui.clearScreen();
        ui.printSectionTitle("Destino de los correos");
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
                    return types.DestInfo{
                        .store_id = try allocator.dupe(u8, ""),
                        .store_name = try allocator.dupe(u8, "Buzon predeterminado"),
                        .store_type = try allocator.dupe(u8, "ExchangeOnline"),
                        .pst_path = try allocator.dupe(u8, ""),
                    };
                } else if (cursor == 1) {
                    const chosen = try store_selector.selectTargetStore(allocator, profile_name);
                    return types.DestInfo{
                        .store_id = chosen.store_id,
                        .store_name = chosen.display_name,
                        .store_type = chosen.store_type,
                        .pst_path = try allocator.dupe(u8, ""),
                    };
                } else {
                    const pst = try selectPstDestination(allocator);
                    return types.DestInfo{
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

fn selectPstDestination(allocator: std.mem.Allocator) ![]u8 {
    var cursor: usize = 0;
    const labels = [_][]const u8{
        "Especificar ruta de archivo PST",
        "Navegar para seleccionar carpeta",
    };

    while (true) {
        ui.clearScreen();
        ui.printSectionTitle("Archivo PST destino");
        std.debug.print("  \x1b[90mPuedes crear un PST nuevo o sobreescribir uno existente.\x1b[0m\n\n", .{});

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
                    return promptPstPath(allocator);
                } else {
                    return browseForPstFolder(allocator);
                }
            },
            else => {},
        }
    }
}

fn promptPstPath(allocator: std.mem.Allocator) ![]u8 {
    ui.clearScreen();
    ui.printSectionTitle("Ruta del archivo PST");
    std.debug.print("  \x1b[90mEjemplo: C:\\Users\\usuario\\Documents\\archivo.pst\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mRuta del PST:\x1b[0m ", .{});

    const path = ui.readLine(allocator) catch return error.Cancelled;
    if (path.len == 0) return error.Cancelled;

    if (std.mem.indexOf(u8, path, ".pst") == null and std.mem.indexOf(u8, path, ".PST") == null) {
        const with_ext = try std.fmt.allocPrint(allocator, "{s}.pst", .{path});
        allocator.free(path);
        return with_ext;
    }

    return path;
}

fn browseForPstFolder(allocator: std.mem.Allocator) ![]u8 {
    const dir = try file_browser.selectPstFile(allocator);
    defer allocator.free(dir);

    ui.clearScreen();
    ui.printSectionTitle("Nombre del archivo PST");
    std.debug.print("  \x1b[90mDirectorio: {s}\x1b[0m\n\n", .{dir});
    std.debug.print("  \x1b[33mNombre del PST (ej: archivo.pst):\x1b[0m ", .{});

    const name = ui.readLine(allocator) catch return error.Cancelled;
    if (name.len == 0) return error.Cancelled;

    const final_name = if (std.mem.indexOf(u8, name, ".pst") != null or std.mem.indexOf(u8, name, ".PST") != null)
        name
    else
        try std.fmt.allocPrint(allocator, "{s}.pst", .{name});

    defer if (final_name.ptr != name.ptr) allocator.free(final_name);

    const full = try std.fs.path.join(allocator, &.{ dir, final_name });
    return full;
}
