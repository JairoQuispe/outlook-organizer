const std = @import("std");
const ui = @import("ui.zig");

const DirEntry = struct {
    name: []const u8,
    is_dir: bool,
    size: u64,
};

/// Interactive file browser that lets user navigate and select a .pst file.
/// Returns the full absolute path to the selected file. Caller owns the memory.
pub fn selectPstFile(allocator: std.mem.Allocator) ![]u8 {
    // Start in user's home or C:\
    var current_path_buf: [std.fs.max_path_bytes]u8 = undefined;
    var current_path: []u8 = undefined;

    const home = std.process.getEnvVarOwned(allocator, "USERPROFILE") catch null;
    if (home) |h| {
        @memcpy(current_path_buf[0..h.len], h);
        current_path = current_path_buf[0..h.len];
        allocator.free(h);
    } else {
        const fallback = "C:\\";
        @memcpy(current_path_buf[0..fallback.len], fallback);
        current_path = current_path_buf[0..fallback.len];
    }

    var browsing_drives = false;

    while (true) {
        // List directory contents
        var entries = std.ArrayListUnmanaged(DirEntry){};
        defer {
            for (entries.items) |entry| {
                allocator.free(entry.name);
            }
            entries.deinit(allocator);
        }

        if (browsing_drives) {
            var letter: u8 = 'A';
            while (letter <= 'Z') : (letter += 1) {
                var drive_buf: [3]u8 = undefined;
                drive_buf[0] = letter;
                drive_buf[1] = ':';
                drive_buf[2] = '\\';

                var test_dir = std.fs.openDirAbsolute(drive_buf[0..], .{}) catch continue;
                test_dir.close();

                const name_copy = allocator.dupe(u8, drive_buf[0..]) catch continue;
                entries.append(allocator, .{
                    .name = name_copy,
                    .is_dir = true,
                    .size = 0,
                }) catch {
                    allocator.free(name_copy);
                };
            }
        } else {
            var dir = std.fs.openDirAbsolute(current_path, .{ .iterate = true }) catch {
                ui.printError("No se puede abrir el directorio");
                ui.waitForEnter();
                // Go to C:\ as fallback
                const fallback = "C:\\";
                @memcpy(current_path_buf[0..fallback.len], fallback);
                current_path = current_path_buf[0..fallback.len];
                browsing_drives = false;
                continue;
            };
            defer dir.close();

            var iter = dir.iterate();
            while (iter.next() catch null) |entry| {
                const is_dir = entry.kind == .directory;
                const is_pst = !is_dir and isPstFile(entry.name);

                if (is_dir or is_pst) {
                    const name_copy = allocator.dupe(u8, entry.name) catch continue;

                    var size: u64 = 0;
                    if (!is_dir) {
                        const stat = dir.statFile(entry.name) catch null;
                        if (stat) |s| {
                            size = s.size;
                        }
                    }

                    entries.append(allocator, .{
                        .name = name_copy,
                        .is_dir = is_dir,
                        .size = size,
                    }) catch {
                        allocator.free(name_copy);
                    };
                }
            }
        }

        // Sort: directories first, then files, both alphabetically
        std.mem.sort(DirEntry, entries.items, {}, struct {
            fn lessThan(_: void, a: DirEntry, b: DirEntry) bool {
                if (a.is_dir and !b.is_dir) return true;
                if (!a.is_dir and b.is_dir) return false;
                return std.mem.lessThan(u8, a.name, b.name);
            }
        }.lessThan);

        var cursor: usize = 0;
        var scroll: usize = 0;
        const page_size: usize = 22;

        browse_loop: while (true) {
            const has_parent_row = !browsing_drives;
            const total_rows = entries.items.len + (if (has_parent_row) @as(usize, 1) else 0);

            if (total_rows == 0) {
                cursor = 0;
                scroll = 0;
            }

            if (total_rows > 0 and cursor >= total_rows) cursor = total_rows - 1;
            if (cursor < scroll) scroll = cursor;
            if (cursor >= scroll + page_size) scroll = cursor - page_size + 1;

            ui.clearScreen();
            ui.printSectionTitle("PST de origen");
            if (browsing_drives) {
                std.debug.print("  \x1b[90mDirectorio: [Unidades del sistema]\x1b[0m\n", .{});
            } else {
                std.debug.print("  \x1b[90mDirectorio: {s}\x1b[0m\n", .{current_path});
            }
            std.debug.print("  \x1b[90m↑/↓ mover | Enter/→ abrir o seleccionar | ←/Backspace subir (raiz => unidades) | P ruta manual | Q cancelar\x1b[0m\n\n", .{});

            const start = scroll;
            const end = @min(total_rows, scroll + page_size);
            var row = start;
            while (row < end) : (row += 1) {
                const is_current = row == cursor;
                if (is_current) std.debug.print("  \x1b[7m", .{});

                if (has_parent_row and row == 0) {
                    std.debug.print("  \x1b[33m[..]\x1b[0m  (Subir)", .{});
                } else {
                    const entry_idx = if (has_parent_row) row - 1 else row;
                    const entry = entries.items[entry_idx];
                    if (entry.is_dir) {
                        if (browsing_drives) {
                            std.debug.print("  \x1b[35m[DRV]\x1b[0m {s}", .{entry.name});
                        } else {
                            std.debug.print("  \x1b[34m[DIR]\x1b[0m {s}", .{entry.name});
                        }
                    } else {
                        const size_mb = @as(f64, @floatFromInt(entry.size)) / (1024.0 * 1024.0);
                        std.debug.print("  \x1b[32m[PST]\x1b[0m {s} \x1b[90m({d:.1} MB)\x1b[0m", .{ entry.name, size_mb });
                    }
                }

                if (is_current) std.debug.print("\x1b[0m", .{});
                std.debug.print("\n", .{});
            }

            if (end < total_rows) {
                std.debug.print("\n  \x1b[90m... y {d} mas\x1b[0m\n", .{total_rows - end});
            }

            const input = ui.readMenuInput(&cursor, total_rows) catch continue;
            switch (input) {
                .cancel => return error.Cancelled,
                .key => |key| switch (key) {
                    'p', 'P' => {
                        if (try handleManualPathInput(allocator, current_path_buf[0..], &current_path, &browsing_drives)) |selected_path| {
                            return selected_path;
                        }
                        break :browse_loop;
                    },
                    8 => {
                        if (browsing_drives) continue;
                        goToParentOrDrives(current_path_buf[0..], &current_path, &browsing_drives);
                        break :browse_loop;
                    },
                    else => {},
                },
                .enter, .right => {
                    var selected_path: ?[]u8 = null;
                    if (try activateSelection(allocator, current_path_buf[0..], &current_path, &browsing_drives, entries.items, has_parent_row, cursor, &selected_path)) {
                        return selected_path.?;
                    }
                    break :browse_loop;
                },
                .left => {
                    goToParentOrDrives(current_path_buf[0..], &current_path, &browsing_drives);
                    break :browse_loop;
                },
                else => {},
            }
        }
    }
}

fn isDriveRoot(path: []const u8) bool {
    if (path.len != 3) return false;
    if (!std.ascii.isAlphabetic(path[0])) return false;
    if (path[1] != ':') return false;
    return path[2] == '\\' or path[2] == '/';
}

fn isPstFile(name: []const u8) bool {
    if (name.len < 5) return false;
    const ext = name[name.len - 4 ..];
    return (std.ascii.toLower(ext[0]) == '.' and
        std.ascii.toLower(ext[1]) == 'p' and
        std.ascii.toLower(ext[2]) == 's' and
        std.ascii.toLower(ext[3]) == 't');
}

fn goToParentOrDrives(current_path_buf: []u8, current_path: *[]u8, browsing_drives: *bool) void {
    if (std.fs.path.dirname(current_path.*)) |parent| {
        @memcpy(current_path_buf[0..parent.len], parent);
        current_path.* = current_path_buf[0..parent.len];
        browsing_drives.* = false;
    } else if (isDriveRoot(current_path.*)) {
        browsing_drives.* = true;
    }
}

fn enterAbsoluteDirectory(current_path_buf: []u8, current_path: *[]u8, browsing_drives: *bool, absolute_path: []const u8) void {
    @memcpy(current_path_buf[0..absolute_path.len], absolute_path);
    current_path.* = current_path_buf[0..absolute_path.len];
    browsing_drives.* = false;
}

fn handleManualPathInput(
    allocator: std.mem.Allocator,
    current_path_buf: []u8,
    current_path: *[]u8,
    browsing_drives: *bool,
) !?[]u8 {
    std.debug.print("\n  \x1b[33mRuta absoluta de archivo PST o carpeta:\x1b[0m ", .{});
    const manual_input = ui.readLine(allocator) catch return null;
    defer allocator.free(manual_input);

    if (manual_input.len == 0) return null;

    if (isPstFile(manual_input)) {
        const stat = std.fs.cwd().statFile(manual_input) catch {
            _ = std.fs.openFileAbsolute(manual_input, .{}) catch {
                ui.printError("Archivo no encontrado");
                ui.waitForEnter();
                return null;
            };
            return try allocator.dupe(u8, manual_input);
        };
        _ = stat;
        return try allocator.dupe(u8, manual_input);
    }

    var test_dir = std.fs.openDirAbsolute(manual_input, .{}) catch {
        ui.printError("Ruta no valida");
        ui.waitForEnter();
        return null;
    };
    test_dir.close();

    enterAbsoluteDirectory(current_path_buf, current_path, browsing_drives, manual_input);
    return null;
}

fn activateSelection(
    allocator: std.mem.Allocator,
    current_path_buf: []u8,
    current_path: *[]u8,
    browsing_drives: *bool,
    entries: []const DirEntry,
    has_parent_row: bool,
    cursor: usize,
    selected_path: *?[]u8,
) !bool {
    if (has_parent_row and cursor == 0) {
        goToParentOrDrives(current_path_buf, current_path, browsing_drives);
        return false;
    }

    const selected_idx = if (has_parent_row) cursor - 1 else cursor;
    const selected = entries[selected_idx];
    if (selected.is_dir) {
        if (browsing_drives.*) {
            enterAbsoluteDirectory(current_path_buf, current_path, browsing_drives, selected.name);
            return false;
        }

        const new_path = std.fs.path.join(allocator, &.{ current_path.*, selected.name }) catch {
            ui.printError("Error construyendo ruta");
            ui.waitForEnter();
            return false;
        };
        defer allocator.free(new_path);
        @memcpy(current_path_buf[0..new_path.len], new_path);
        current_path.* = current_path_buf[0..new_path.len];
        browsing_drives.* = false;
        return false;
    }

    selected_path.* = try std.fs.path.join(allocator, &.{ current_path.*, selected.name });
    return true;
}
