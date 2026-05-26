const std = @import("std");
const utils = @import("utils.zig");

pub const EventType = enum {
    progress,
    folder,
    err,
    restore_result,
    other,
};

pub const ParsedEvent = struct {
    event_type: EventType,
    line: []const u8,
    copied: ?i64 = null,
    moved: ?i64 = null,
    skipped: ?i64 = null,
    failed: ?i64 = null,
    size_bytes: ?i64 = null,
    percent: ?i64 = null,
    status: ?[]const u8 = null,
    message: ?[]const u8 = null,
    path: ?[]const u8 = null,
    item_count: ?i64 = null,
};

pub fn parseEventLine(line: []const u8) ?ParsedEvent {
    if (line.len == 0) return null;
    const kind = utils.extractString(line, "type") orelse return null;

    if (std.mem.eql(u8, kind, "progress")) {
        return .{
            .event_type = .progress,
            .line = line,
            .copied = utils.extractNumber(line, "copied"),
            .moved = utils.extractNumber(line, "moved"),
            .skipped = utils.extractNumber(line, "skipped"),
            .failed = utils.extractNumber(line, "failed"),
            .size_bytes = utils.extractNumber(line, "sizeBytes"),
            .percent = utils.extractNumber(line, "percent"),
            .status = utils.extractString(line, "status"),
        };
    }

    if (std.mem.eql(u8, kind, "folder")) {
        return .{
            .event_type = .folder,
            .line = line,
            .path = utils.extractString(line, "path"),
            .item_count = utils.extractNumber(line, "itemCount"),
            .size_bytes = utils.extractNumber(line, "sizeBytes"),
        };
    }

    if (std.mem.eql(u8, kind, "error")) {
        return .{
            .event_type = .err,
            .line = line,
            .message = utils.extractString(line, "message"),
        };
    }

    if (std.mem.eql(u8, kind, "restoreResult")) {
        return .{
            .event_type = .restore_result,
            .line = line,
            .copied = utils.extractNumber(line, "copied"),
            .moved = utils.extractNumber(line, "moved"),
            .skipped = utils.extractNumber(line, "skipped"),
            .failed = utils.extractNumber(line, "failed"),
        };
    }

    return .{
        .event_type = .other,
        .line = line,
    };
}
