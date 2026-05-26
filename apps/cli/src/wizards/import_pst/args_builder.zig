const std = @import("std");

pub const ArgsBuilder = struct {
    allocator: std.mem.Allocator,
    list: std.ArrayListUnmanaged([]const u8),

    pub fn init(allocator: std.mem.Allocator) ArgsBuilder {
        return .{
            .allocator = allocator,
            .list = .{},
        };
    }

    pub fn deinit(self: *ArgsBuilder) void {
        self.list.deinit(self.allocator);
    }

    pub fn items(self: *const ArgsBuilder) []const []const u8 {
        return self.list.items;
    }

    pub fn addSlice(self: *ArgsBuilder, values: []const []const u8) !void {
        try self.list.appendSlice(self.allocator, values);
    }

    pub fn addFlag(self: *ArgsBuilder, flag: []const u8) !void {
        try self.list.append(self.allocator, flag);
    }

    pub fn addOption(self: *ArgsBuilder, flag: []const u8, value: []const u8) !void {
        try self.list.append(self.allocator, flag);
        try self.list.append(self.allocator, value);
    }

    pub fn addBoolFlag(self: *ArgsBuilder, enabled: bool, flag: []const u8) !void {
        if (!enabled) return;
        try self.addFlag(flag);
    }

    pub fn addOptionIfNonEmpty(self: *ArgsBuilder, flag: []const u8, value: ?[]const u8) !void {
        const v = value orelse return;
        if (v.len == 0) return;
        try self.addOption(flag, v);
    }
};
