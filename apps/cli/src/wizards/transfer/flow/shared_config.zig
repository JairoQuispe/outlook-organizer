const std = @import("std");
const types = @import("../types.zig");

pub const SharedConfig = struct {
    profile_name: ?[]const u8,
    enable_routing: bool,
    routing_criterion: ?types.RoutingCriterion,
    scan_mode: types.ScanMode,
    scan_filter_year: ?[]const u8,
    action: []const u8,
    skip_duplicates: bool,
    deep_duplicate_check: bool,
    filter_year: ?[]const u8,
    filter_months: ?[]const u8,
    adaptive_throttling: bool,

    pub fn initDefaults() SharedConfig {
        return .{
            .profile_name = null,
            .enable_routing = false,
            .routing_criterion = null,
            .scan_mode = .quick,
            .scan_filter_year = null,
            .action = "Copy",
            .skip_duplicates = true,
            .deep_duplicate_check = false,
            .filter_year = null,
            .filter_months = null,
            .adaptive_throttling = true,
        };
    }

    pub fn deinit(self: *SharedConfig, allocator: std.mem.Allocator) void {
        if (self.profile_name) |p| allocator.free(p);
        if (self.scan_filter_year) |y| allocator.free(y);
        if (self.filter_year) |y| allocator.free(y);
        if (self.filter_months) |m| allocator.free(m);
    }
};
