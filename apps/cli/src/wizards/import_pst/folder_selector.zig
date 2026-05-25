const std = @import("std");
const ui = @import("../../ui.zig");
const types = @import("types.zig");
const utils = @import("utils.zig");

const ScannedFolder = types.ScannedFolder;
const FolderTree = types.FolderTree;
const FolderTreeNode = types.FolderTreeNode;
const VisibleTreeRow = types.VisibleTreeRow;

pub fn promptFolderSelection(allocator: std.mem.Allocator, folders: []const ScannedFolder, scan_mode: types.ScanMode) ![]bool {
    var tree = try buildFolderTree(allocator, folders);
    defer deinitFolderTree(allocator, &tree);

    var selected = try allocator.alloc(bool, folders.len);
    @memset(selected, false);

    var cursor: usize = 0;
    var scroll: usize = 0;
    const page_size: usize = 22;

    while (true) {
        var visible = std.ArrayListUnmanaged(VisibleTreeRow){};
        defer visible.deinit(allocator);
        try buildVisibleRows(allocator, &tree, &visible);

        if (visible.items.len == 0) return selected;
        if (cursor >= visible.items.len) cursor = visible.items.len - 1;

        if (cursor < scroll) scroll = cursor;
        if (cursor >= scroll + page_size) scroll = cursor - page_size + 1;

        ui.clearScreen();
        ui.printSectionTitle("Seleccionar carpetas");
        std.debug.print("  \x1b[90mEspacio: marcar | A: todas | C: contraer | E: expandir | W/S o ↑↓: mover | ←: contraer/desmarcar | →: expandir/marcar | Enter: confirmar | Q: cancelar\x1b[0m\n", .{});
        std.debug.print("  \x1b[90mModo: {s} | Seleccionadas: {d}/{d}\x1b[0m\n\n", .{ if (scan_mode == .deep) "Profundo" else "Rapido", countSelectedFlags(selected), folders.len });

        const start = scroll;
        const end = @min(visible.items.len, scroll + page_size);
        for (visible.items[start..end], start..) |row, row_index| {
            const node = tree.nodes.items[row.node_index];
            const is_current = row_index == cursor;
            const is_selected = if (node.folder_index) |fi| selected[fi] else false;
            const has_children = node.children.items.len > 0;
            const branch = if (has_children) (if (node.expanded) "[-]" else "[+]") else "   ";
            const mark = if (is_selected) "x" else " ";

            if (is_current) std.debug.print("  \x1b[7m", .{});
            std.debug.print("  ", .{});
            var d: usize = 0;
            while (d < row.depth) : (d += 1) {
                std.debug.print("  ", .{});
            }
            std.debug.print("{s} [{s}] {s}", .{ branch, mark, node.name });

            if (node.folder_index) |fi| {
                const folder = folders[fi];
                if (scan_mode == .deep and folder.size_bytes != null) {
                    const size_mb = @as(f64, @floatFromInt(folder.size_bytes.?)) / (1024.0 * 1024.0);
                    if (folder.year_breakdown_display.len > 0) {
                        std.debug.print(" \x1b[90m(items:{d}, anios:{s}, {d:.1}MB)\x1b[0m", .{ folder.item_count, folder.year_breakdown_display, size_mb });
                    } else {
                        std.debug.print(" \x1b[90m(items:{d}, {d:.1}MB)\x1b[0m", .{ folder.item_count, size_mb });
                    }
                } else {
                    if (folder.year_breakdown_display.len > 0) {
                        std.debug.print(" \x1b[90m(items:{d}, anios:{s})\x1b[0m", .{ folder.item_count, folder.year_breakdown_display });
                    } else {
                        std.debug.print(" \x1b[90m(items:{d})\x1b[0m", .{folder.item_count});
                    }
                }
            }
            if (is_current) std.debug.print("\x1b[0m", .{});
            std.debug.print("\n", .{});
        }

        if (end < visible.items.len) {
            std.debug.print("\n  \x1b[90m... {d} carpetas mas\x1b[0m\n", .{visible.items.len - end});
        }

        const key = ui.readSingleKey() catch continue;
        switch (key) {
            'q', 'Q' => {
                allocator.free(selected);
                return error.Cancelled;
            },
            'a', 'A' => {
                @memset(selected, true);
            },
            'c', 'C' => collapseAllNodes(&tree),
            'e', 'E' => expandAllNodes(&tree),
            'w', 'W', 'k', 'K' => {
                if (cursor > 0) cursor -= 1;
            },
            's', 'S', 'j', 'J' => {
                if (cursor + 1 < visible.items.len) cursor += 1;
            },
            ' ' => {
                const node = tree.nodes.items[visible.items[cursor].node_index];
                if (node.folder_index) |fi| {
                    selected[fi] = !selected[fi];
                }
            },
            '\r', '\n' => {
                if (countSelectedFlags(selected) == 0) continue;
                return selected;
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
                        if (cursor + 1 < visible.items.len) cursor += 1;
                    },
                    'D' => {
                        const node_index = visible.items[cursor].node_index;
                        if (tree.nodes.items[node_index].children.items.len > 0) {
                            tree.nodes.items[node_index].expanded = false;
                        } else if (tree.nodes.items[node_index].folder_index) |fi| {
                            selected[fi] = false;
                        }
                    },
                    'C' => {
                        const node_index = visible.items[cursor].node_index;
                        if (tree.nodes.items[node_index].children.items.len > 0) {
                            tree.nodes.items[node_index].expanded = true;
                        } else if (tree.nodes.items[node_index].folder_index) |fi| {
                            selected[fi] = true;
                        }
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
                        if (cursor + 1 < visible.items.len) cursor += 1;
                    },
                    75 => {
                        const node_index = visible.items[cursor].node_index;
                        if (tree.nodes.items[node_index].children.items.len > 0) {
                            tree.nodes.items[node_index].expanded = false;
                        } else if (tree.nodes.items[node_index].folder_index) |fi| {
                            selected[fi] = false;
                        }
                    },
                    77 => {
                        const node_index = visible.items[cursor].node_index;
                        if (tree.nodes.items[node_index].children.items.len > 0) {
                            tree.nodes.items[node_index].expanded = true;
                        } else if (tree.nodes.items[node_index].folder_index) |fi| {
                            selected[fi] = true;
                        }
                    },
                    else => {},
                }
            },
            else => {},
        }
    }
}

pub fn buildFolderTree(allocator: std.mem.Allocator, folders: []const ScannedFolder) !FolderTree {
    var tree = FolderTree{
        .nodes = .{},
        .roots = .{},
    };
    errdefer deinitFolderTree(allocator, &tree);

    for (folders, 0..) |folder, folder_index| {
        var parent: ?usize = null;
        var iter = std.mem.splitScalar(u8, folder.path, '\\');
        var prefix = std.ArrayListUnmanaged(u8){};
        defer prefix.deinit(allocator);

        while (iter.next()) |part| {
            const name = std.mem.trim(u8, part, &[_]u8{' '});
            if (name.len == 0) continue;

            if (prefix.items.len > 0) try prefix.append(allocator, '\\');
            try prefix.appendSlice(allocator, name);

            var found = findChildByName(&tree, parent, name);
            if (found == null) {
                const node_name = try allocator.dupe(u8, name);
                errdefer allocator.free(node_name);
                const node_path = try allocator.dupe(u8, prefix.items);
                errdefer allocator.free(node_path);

                const new_index = tree.nodes.items.len;
                try tree.nodes.append(allocator, .{
                    .name = node_name,
                    .full_path = node_path,
                    .parent = parent,
                    .children = .{},
                    .expanded = true,
                    .folder_index = null,
                });

                if (parent) |p| {
                    try tree.nodes.items[p].children.append(allocator, new_index);
                } else {
                    try tree.roots.append(allocator, new_index);
                }
                found = new_index;
            }

            parent = found;
        }

        if (parent) |idx| {
            tree.nodes.items[idx].folder_index = folder_index;
        }
    }

    return tree;
}

pub fn deinitFolderTree(allocator: std.mem.Allocator, tree: *FolderTree) void {
    for (tree.nodes.items) |*node| {
        allocator.free(node.name);
        allocator.free(node.full_path);
        node.children.deinit(allocator);
    }
    tree.nodes.deinit(allocator);
    tree.roots.deinit(allocator);
}

pub fn findChildByName(tree: *const FolderTree, parent: ?usize, name: []const u8) ?usize {
    if (parent) |p| {
        for (tree.nodes.items[p].children.items) |child_idx| {
            if (std.mem.eql(u8, tree.nodes.items[child_idx].name, name)) return child_idx;
        }
        return null;
    }

    for (tree.roots.items) |root_idx| {
        if (std.mem.eql(u8, tree.nodes.items[root_idx].name, name)) return root_idx;
    }
    return null;
}

pub fn buildVisibleRows(allocator: std.mem.Allocator, tree: *const FolderTree, out: *std.ArrayListUnmanaged(VisibleTreeRow)) !void {
    for (tree.roots.items) |root_idx| {
        try appendVisibleRowsRecursive(allocator, tree, root_idx, 0, out);
    }
}

pub fn appendVisibleRowsRecursive(
    allocator: std.mem.Allocator,
    tree: *const FolderTree,
    node_index: usize,
    depth: usize,
    out: *std.ArrayListUnmanaged(VisibleTreeRow),
) !void {
    try out.append(allocator, .{ .node_index = node_index, .depth = depth });
    const node = tree.nodes.items[node_index];
    if (!node.expanded) return;

    for (node.children.items) |child_idx| {
        try appendVisibleRowsRecursive(allocator, tree, child_idx, depth + 1, out);
    }
}

pub fn collapseAllNodes(tree: *FolderTree) void {
    for (tree.nodes.items) |*node| {
        if (node.children.items.len > 0) node.expanded = false;
    }
}

pub fn expandAllNodes(tree: *FolderTree) void {
    for (tree.nodes.items) |*node| {
        if (node.children.items.len > 0) node.expanded = true;
    }
}

pub fn countSelectedFlags(flags: []const bool) usize {
    var count: usize = 0;
    for (flags) |flag| {
        if (flag) count += 1;
    }
    return count;
}

pub fn writeFolderPlanFromFlags(allocator: std.mem.Allocator, folders: []const ScannedFolder, selected_flags: []const bool) ![]u8 {
    const plan_path = try utils.makeTempFilePath(allocator, "oo-folder-plan", "json");
    errdefer allocator.free(plan_path);

    var json = std.ArrayListUnmanaged(u8){};
    defer json.deinit(allocator);

    try json.appendSlice(allocator, "{\"type\":\"folderExport\",\"folders\":[");

    var first = true;
    for (folders, 0..) |folder, idx| {
        if (!selected_flags[idx]) continue;
        if (!first) try json.appendSlice(allocator, ",");
        first = false;

        try json.appendSlice(allocator, "{\"path\":\"");
        try utils.appendJsonEscaped(allocator, &json, folder.path);
        try json.appendSlice(allocator, "\",\"itemCount\":");
        const num_str = try std.fmt.allocPrint(allocator, "{d}", .{folder.item_count});
        defer allocator.free(num_str);
        try json.appendSlice(allocator, num_str);
        try json.appendSlice(allocator, "}");
    }

    try json.appendSlice(allocator, "]}");

    const file = std.fs.createFileAbsolute(plan_path, .{ .truncate = true }) catch {
        allocator.free(plan_path);
        return error.CannotWriteFolderPlan;
    };
    defer file.close();

    try file.writeAll(json.items);
    return plan_path;
}
