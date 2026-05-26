const std = @import("std");
const types = @import("../types.zig");
const utils = @import("../utils.zig");
const event_parser = @import("../event_parser.zig");

pub fn parseScannedFolders(allocator: std.mem.Allocator, output: []const u8) !std.ArrayListUnmanaged(types.ScannedFolder) {
    var folders = std.ArrayListUnmanaged(types.ScannedFolder){};
    errdefer {
        for (folders.items) |folder| {
            allocator.free(folder.path);
            allocator.free(folder.year_breakdown_display);
        }
        folders.deinit(allocator);
    }

    var line_iter = std.mem.splitSequence(u8, output, "\n");
    while (line_iter.next()) |raw_line| {
        const line = std.mem.trim(u8, raw_line, &[_]u8{ ' ', '\t', '\r', '\n' });
        const evt = event_parser.parseEventLine(line) orelse continue;
        if (evt.event_type != .folder) continue;

        const item_count = evt.item_count orelse 0;
        if (item_count <= 0) continue;

        const path = evt.path orelse continue;
        const path_copy = try allocator.dupe(u8, path);
        errdefer allocator.free(path_copy);

        const size_bytes = evt.size_bytes;
        const year_breakdown_display = try utils.extractYearBreakdownDisplay(allocator, line);
        errdefer allocator.free(year_breakdown_display);

        try folders.append(allocator, .{
            .path = path_copy,
            .item_count = item_count,
            .size_bytes = size_bytes,
            .year_breakdown_display = year_breakdown_display,
        });
    }

    return folders;
}
