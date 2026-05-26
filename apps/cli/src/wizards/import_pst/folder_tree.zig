const std = @import("std");
const types = @import("types.zig");

const ScannedFolder = types.ScannedFolder;
const FolderTree = types.FolderTree;
const FolderTreeNode = types.FolderTreeNode;
const VisibleTreeRow = types.VisibleTreeRow;

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
