const std = @import("std");
const ui = @import("../../ui.zig");

pub fn beginMenu(title: []const u8) void {
    ui.clearScreen();
    ui.printSectionTitle(title);
}

pub fn printContext(label: []const u8, value: []const u8) void {
    std.debug.print("  \x1b[90m{s}:\x1b[0m {s}\n", .{ label, value });
}

pub fn printMutedLine(line: []const u8) void {
    std.debug.print("  \x1b[90m{s}\x1b[0m\n", .{line});
}

pub fn beginHighlightedRow(is_current: bool) void {
    if (is_current) std.debug.print("  \x1b[7m", .{});
}

pub fn endHighlightedRow(is_current: bool) void {
    if (is_current) std.debug.print("  \x1b[0m", .{});
}

pub fn printSelectableLabel(label: []const u8, is_current: bool) void {
    beginHighlightedRow(is_current);
    std.debug.print("  {s}\n", .{label});
    endHighlightedRow(is_current);
}
