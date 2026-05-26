const std = @import("std");
const types = @import("../types.zig");
const utils = @import("../utils.zig");

pub const BatchPstConfig = struct {
    pst_path: []const u8,
    folder_plan_path: []const u8,
    scan_export_path: []const u8,
    selected_count: usize,
    total_count: usize,
    scan_mode: types.ScanMode,
    target_store_id: []const u8,
    target_store_name: []const u8,
    target_store_type: []const u8,
    routing_mappings: ?[]const types.TargetStoreMapping,

    pub fn deinit(self: *BatchPstConfig, allocator: std.mem.Allocator) void {
        utils.cleanupTempFile(self.folder_plan_path);
        allocator.free(self.folder_plan_path);
        utils.cleanupTempFile(self.scan_export_path);
        allocator.free(self.scan_export_path);
        allocator.free(self.target_store_id);
        allocator.free(self.target_store_name);
        allocator.free(self.target_store_type);
        if (self.routing_mappings) |mappings| {
            utils.freeTargetStoreMappings(allocator, mappings);
        }
    }
};
