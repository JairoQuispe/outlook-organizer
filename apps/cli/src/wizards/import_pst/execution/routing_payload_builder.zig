const std = @import("std");
const types = @import("../types.zig");
const utils = @import("../utils.zig");

pub fn buildRoutingMappingsJson(allocator: std.mem.Allocator, mappings: []const types.TargetStoreMapping) ![]u8 {
    var buf = std.ArrayListUnmanaged(u8){};
    defer buf.deinit(allocator);

    try buf.writer(allocator).writeAll("[");
    var wrote_any = false;
    for (mappings) |m| {
        if (m.store_id.len == 0) continue;
        if (wrote_any) try buf.writer(allocator).writeAll(",");
        try buf.writer(allocator).print("{{\"year\":{d}", .{m.year});
        if (m.month) |mon| {
            try buf.writer(allocator).print(",\"month\":{d}", .{mon});
        } else {
            try buf.writer(allocator).writeAll(",\"month\":null");
        }
        try buf.writer(allocator).writeAll(",\"storeId\":\"");
        try utils.appendJsonEscaped(allocator, &buf, m.store_id);
        try buf.writer(allocator).writeAll("\"}");
        wrote_any = true;
    }
    try buf.writer(allocator).writeAll("]");

    if (!wrote_any) return error.NoValidMappings;
    return try buf.toOwnedSlice(allocator);
}
