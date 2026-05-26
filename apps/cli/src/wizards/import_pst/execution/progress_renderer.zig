const std = @import("std");
const ui = @import("../../../ui.zig");
const types = @import("../types.zig");
const utils = @import("../utils.zig");
const event_parser = @import("../event_parser.zig");

const ParsedProgress = struct {
    processed: i64,
    total: i64,
};

fn parseProcessedTotalFromStatus(status: []const u8) ?ParsedProgress {
    const close_idx = std.mem.lastIndexOfScalar(u8, status, ')') orelse return null;
    const open_idx = std.mem.lastIndexOfScalar(u8, status[0..close_idx], '(') orelse return null;
    if (open_idx >= close_idx) return null;

    const inner = std.mem.trim(u8, status[open_idx + 1 .. close_idx], " ");
    const slash_idx = std.mem.indexOfScalar(u8, inner, '/') orelse return null;
    const left = std.mem.trim(u8, inner[0..slash_idx], " ");
    const right = std.mem.trim(u8, inner[slash_idx + 1 ..], " ");
    if (left.len == 0 or right.len == 0) return null;

    const processed = std.fmt.parseInt(i64, left, 10) catch return null;
    const total = std.fmt.parseInt(i64, right, 10) catch return null;
    if (processed < 0 or total <= 0) return null;

    return .{ .processed = processed, .total = total };
}

pub fn onImportScriptLine(ctx: *anyopaque, line: []const u8) void {
    var state: *types.ImportProgressState = @ptrCast(@alignCast(ctx));
    const evt = event_parser.parseEventLine(line) orelse return;
    if (evt.event_type != .progress) return;

    if (evt.copied) |v| state.copied = v;
    if (evt.moved) |v| state.moved = v;
    if (evt.skipped) |v| state.skipped = v;
    if (evt.failed) |v| state.failed = v;
    if (evt.size_bytes) |v| state.size_bytes = v;
    if (evt.percent) |v| {
        if (v <= 0) {
            state.percent = 0;
        } else if (v >= 100) {
            state.percent = 100;
        } else {
            state.percent = @intCast(v);
        }
    }

    const total_script_processed = state.copied + state.moved + state.skipped + state.failed;
    const elapsed_ms = std.time.milliTimestamp() - state.start_ms;

    if (state.has_rendered_progress) {
        std.debug.print("\x1b[7F", .{});
    } else {
        state.has_rendered_progress = true;
    }

    const current_status = evt.status orelse "";
    const parsed = parseProcessedTotalFromStatus(current_status);

    const processed_effective = if (parsed) |p|
        @max(p.processed, total_script_processed)
    else
        total_script_processed;

    const remaining_effective = if (parsed) |p|
        @max(p.total - processed_effective, 0)
    else blk: {
        const percent_fallback: u32 = if (state.percent > 0)
            if (state.percent >= 100) 100 else state.percent
        else
            @as(u32, if (processed_effective > 0) 1 else 0);

        var estimated_total = processed_effective;
        if (percent_fallback > 0) {
            var est = @divTrunc(processed_effective * 100, @as(i64, @intCast(percent_fallback)));
            if (est < processed_effective) est = processed_effective;
            estimated_total = est;
        }

        break :blk @max(estimated_total - processed_effective, 0);
    };

    const percent_effective: u32 = if (parsed) |p| blk: {
        const capped_processed = @min(processed_effective, p.total);
        var pct = @as(u32, @intCast(@divTrunc(capped_processed * 100, p.total)));
        if (pct > 99 and capped_processed < p.total) pct = 99;
        break :blk pct;
    } else if (state.percent > 0)
        if (state.percent >= 100) 100 else state.percent
    else
        @as(u32, if (processed_effective > 0) 1 else 0);

    const columns = ui.terminalWidthColumns();
    const status_prefix_len = "  Carpeta: ".len;
    const status_max = if (columns > status_prefix_len) columns - status_prefix_len else 0;
    var status_buf: [512]u8 = undefined;
    const safe_status = ui.truncateWithEllipsis(current_status, &status_buf, status_max);

    std.debug.print("\r\x1b[2K  \x1b[90mCarpeta:\x1b[0m {s}\n", .{safe_status});

    ui.printProgressBar(percent_effective, "Importando");

    var size_buf: [32]u8 = undefined;
    const size_str = utils.formatBytesShort(&size_buf, state.size_bytes);

    std.debug.print("\n\x1b[2K  \x1b[90mProc:\x1b[0m {d} ({s})", .{ processed_effective, size_str });
    std.debug.print("\n\x1b[2K  \x1b[90mCop:\x1b[0m  {d}", .{state.copied});
    std.debug.print("\n\x1b[2K  \x1b[90mMov:\x1b[0m  {d}", .{state.moved});
    std.debug.print("\n\x1b[2K  \x1b[90mOmi:\x1b[0m  {d}", .{state.skipped + state.failed});
    std.debug.print("\n\x1b[2K  \x1b[90mRes:\x1b[0m  {d}", .{remaining_effective});
    var elapsed_hms_buf: [32]u8 = undefined;
    const elapsed_hms = utils.formatHms(&elapsed_hms_buf, @divTrunc(elapsed_ms, 1000));
    std.debug.print("\n\x1b[2K  \x1b[90mT:\x1b[0m {s}", .{elapsed_hms});

    var eta_str_buf: [32]u8 = undefined;
    var eta_str: []const u8 = "--:--:--";
    if (percent_effective > 0 and percent_effective < 100) {
        const total_est_ms = @divTrunc(elapsed_ms * 100, @as(i64, @intCast(percent_effective)));
        const eta_ms = @max(total_est_ms - elapsed_ms, 0);
        const eta_s = @divTrunc(eta_ms, 1000);
        eta_str = utils.formatHms(&eta_str_buf, eta_s);
    } else if (percent_effective >= 100) {
        eta_str = "00:00:00";
    }
    std.debug.print("  \x1b[90mETA:\x1b[0m {s}", .{eta_str});
    std.debug.print("\r", .{});
}
