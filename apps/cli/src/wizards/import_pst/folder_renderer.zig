const std = @import("std");
const types = @import("types.zig");
const utils = @import("utils.zig");

const ScannedFolder = types.ScannedFolder;
const FolderTreeNode = types.FolderTreeNode;
const VisibleTreeRow = types.VisibleTreeRow;

fn printFolderInfo(folder: *const ScannedFolder, scan_mode: types.ScanMode) void {
    if (scan_mode == .deep and folder.size_bytes != null) {
        var size_buf: [32]u8 = undefined;
        const size_str = utils.formatBytesShort(&size_buf, folder.size_bytes.?);
        if (folder.year_breakdown_display.len > 0) {
            std.debug.print(" \x1b[90m(items:{d}, anios:{s}, {s})\x1b[0m", .{ folder.item_count, folder.year_breakdown_display, size_str });
        } else {
            std.debug.print(" \x1b[90m(items:{d}, {s})\x1b[0m", .{ folder.item_count, size_str });
        }
    } else {
        if (folder.year_breakdown_display.len > 0) {
            std.debug.print(" \x1b[90m(items:{d}, anios:{s})\x1b[0m", .{ folder.item_count, folder.year_breakdown_display });
        } else {
            std.debug.print(" \x1b[90m(items:{d})\x1b[0m", .{folder.item_count});
        }
    }
}

pub fn renderTreeRow(row: VisibleTreeRow, node: *const FolderTreeNode, folders: []const ScannedFolder, selected: []const bool, cursor: usize, row_index: usize, scan_mode: types.ScanMode) void {
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
        printFolderInfo(&folders[fi], scan_mode);
    }
    if (is_current) std.debug.print("\x1b[0m", .{});
    std.debug.print("\n", .{});
}
