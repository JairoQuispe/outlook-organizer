const std = @import("std");
const ui = @import("../../ui.zig");
const types = @import("types.zig");
const utils = @import("utils.zig");
const folder_tree = @import("folder_tree.zig");
const folder_renderer = @import("folder_renderer.zig");

const ScannedFolder = types.ScannedFolder;

pub fn promptFolderSelection(allocator: std.mem.Allocator, folders: []const ScannedFolder, scan_mode: types.ScanMode) ![]bool {
    var tree = try folder_tree.buildFolderTree(allocator, folders);
    defer folder_tree.deinitFolderTree(allocator, &tree);

    var selected = try allocator.alloc(bool, folders.len);
    @memset(selected, false);

    var cursor: usize = 0;
    var scroll: usize = 0;
    const page_size: usize = 22;

    while (true) {
        var visible = std.ArrayListUnmanaged(types.VisibleTreeRow){};
        defer visible.deinit(allocator);
        try folder_tree.buildVisibleRows(allocator, &tree, &visible);

        if (visible.items.len == 0) return selected;
        if (cursor >= visible.items.len) cursor = visible.items.len - 1;

        if (cursor < scroll) scroll = cursor;
        if (cursor >= scroll + page_size) scroll = cursor - page_size + 1;

        ui.clearScreen();
        ui.printSectionTitle("Seleccionar carpetas");
        std.debug.print("  \x1b[90mEspacio: marcar | A: todas | C: contraer | E: expandir | W/S o ↑↓: mover | ←: contraer/desmarcar | →: expandir/marcar | Enter: confirmar | Q: cancelar\x1b[0m\n", .{});
        std.debug.print("  \x1b[90mModo: {s} | Seleccionadas: {d}/{d}\x1b[0m\n\n", .{ if (scan_mode == .deep) "Profundo" else "Rapido", @max(countSelectedFlags(selected), 0), folders.len });

        const start = scroll;
        const end = @min(visible.items.len, scroll + page_size);
        for (visible.items[start..end], start..) |row, row_index| {
            const node = tree.nodes.items[row.node_index];
            folder_renderer.renderTreeRow(row, &node, folders, selected, cursor, row_index, scan_mode);
        }

        if (end < visible.items.len) {
            std.debug.print("\n  \x1b[90m... {d} carpetas mas\x1b[0m\n", .{visible.items.len - end});
        }

        const input = ui.readMenuInput(&cursor, visible.items.len) catch continue;
        switch (input) {
            .cancel => {
                allocator.free(selected);
                return error.Cancelled;
            },
            .enter => {
                if (countSelectedFlags(selected) == 0) continue;
                return selected;
            },
            .left => {
                const node_index = visible.items[cursor].node_index;
                if (tree.nodes.items[node_index].children.items.len > 0) {
                    tree.nodes.items[node_index].expanded = false;
                } else if (tree.nodes.items[node_index].folder_index) |fi| {
                    selected[fi] = false;
                }
            },
            .right => {
                const node_index = visible.items[cursor].node_index;
                if (tree.nodes.items[node_index].children.items.len > 0) {
                    tree.nodes.items[node_index].expanded = true;
                } else if (tree.nodes.items[node_index].folder_index) |fi| {
                    selected[fi] = true;
                }
            },
            .key => |key| switch (key) {
                'a', 'A' => @memset(selected, true),
                'c', 'C' => folder_tree.collapseAllNodes(&tree),
                'e', 'E' => folder_tree.expandAllNodes(&tree),
                ' ' => {
                    const node = tree.nodes.items[visible.items[cursor].node_index];
                    if (node.folder_index) |fi| {
                        selected[fi] = !selected[fi];
                    }
                },
                else => {},
            },
            else => {},
        }
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

pub fn cleanupFolderSelection(allocator: std.mem.Allocator, result: *const types.FolderSelectionResult) void {
    utils.cleanupTempFile(result.folder_plan_path);
    allocator.free(result.folder_plan_path);
    utils.cleanupTempFile(result.scan_export_path);
    allocator.free(result.scan_export_path);
}
