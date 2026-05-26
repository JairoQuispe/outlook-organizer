const std = @import("std");
const event_parser = @import("../event_parser.zig");

pub const ParsedImportOutput = struct {
    last_copied: i64 = 0,
    last_moved: i64 = 0,
    last_skipped: i64 = 0,
    last_failed: i64 = 0,
    result_line: ?[]const u8 = null,
    error_message: ?[]const u8 = null,
};

pub fn parseImportOutput(output: []const u8) ParsedImportOutput {
    var parsed = ParsedImportOutput{};
    var line_iter = std.mem.splitSequence(u8, output, "\n");
    while (line_iter.next()) |raw_line| {
        const line = std.mem.trim(u8, raw_line, &[_]u8{ ' ', '\t', '\r', '\n' });
        if (line.len == 0) continue;
        const evt = event_parser.parseEventLine(line) orelse continue;

        switch (evt.event_type) {
            .progress => {
                if (evt.copied) |c| parsed.last_copied = c;
                if (evt.moved) |m| parsed.last_moved = m;
                if (evt.skipped) |s| parsed.last_skipped = s;
                if (evt.failed) |f| parsed.last_failed = f;
            },
            .err => {
                parsed.error_message = evt.message orelse line;
            },
            .restore_result => {
                parsed.result_line = line;
            },
            else => {},
        }
    }
    return parsed;
}
