const std = @import("std");
const builtin = @import("builtin");
const windows = std.os.windows;

var ansi_initialized = false;
var key_pending: [16]u8 = undefined;
var key_pending_pos: usize = 0;
var key_pending_len: usize = 0;

pub fn ensureAnsiMode() void {
    if (builtin.os.tag != .windows or ansi_initialized) return;
    ansi_initialized = true;

    const handle = windows.GetStdHandle(windows.STD_OUTPUT_HANDLE) catch return;

    var mode: windows.DWORD = undefined;
    if (windows.kernel32.GetConsoleMode(handle, &mode) == 0) return;

    const ENABLE_VIRTUAL_TERMINAL_PROCESSING: windows.DWORD = 0x0004;
    if ((mode & ENABLE_VIRTUAL_TERMINAL_PROCESSING) != 0) return;

    _ = windows.kernel32.SetConsoleMode(handle, mode | ENABLE_VIRTUAL_TERMINAL_PROCESSING);
}

pub const MenuChoice = enum {
    import_pst,
    scan_pst,
    exit,
};

pub const MenuInput = union(enum) {
    up,
    down,
    left,
    right,
    enter,
    cancel,
    key: u8,
};

pub fn readMenuInput(cursor: *usize, item_count: usize) !MenuInput {
    const input = try readSingleKey();

    const nav = switch (input) {
        'q', 'Q' => MenuInput{ .cancel = {} },
        'w', 'W', 'k', 'K' => MenuInput{ .up = {} },
        's', 'S', 'j', 'J' => MenuInput{ .down = {} },
        '\r', '\n' => MenuInput{ .enter = {} },
        27 => blk: {
            const seq1 = readSingleKey() catch return MenuInput{ .key = input };
            if (seq1 != '[' and seq1 != 'O') return MenuInput{ .key = input };

            const seq2 = readSingleKey() catch return MenuInput{ .key = input };
            break :blk switch (seq2) {
                'A' => MenuInput{ .up = {} },
                'B' => MenuInput{ .down = {} },
                'C' => MenuInput{ .right = {} },
                'D' => MenuInput{ .left = {} },
                else => MenuInput{ .key = input },
            };
        },
        0, 224 => blk: {
            const ext = readSingleKey() catch return MenuInput{ .key = input };
            break :blk switch (ext) {
                72 => MenuInput{ .up = {} },
                80 => MenuInput{ .down = {} },
                77 => MenuInput{ .right = {} },
                75 => MenuInput{ .left = {} },
                else => MenuInput{ .key = input },
            };
        },
        else => MenuInput{ .key = input },
    };

    switch (nav) {
        .up => {
            if (cursor.* > 0) cursor.* -= 1;
        },
        .down => {
            if (cursor.* + 1 < item_count) cursor.* += 1;
        },
        else => {},
    }

    return nav;
}

pub fn clearScreen() void {
    // ANSI escape: clear screen + move cursor to top-left
    std.debug.print("\x1b[2J\x1b[H", .{});
}

pub fn printBanner() void {
    std.debug.print("\n", .{});
    std.debug.print("  \x1b[36m   ____        __  __            __             \x1b[0m\n", .{});
    std.debug.print("  \x1b[36m  / __ \\__  __/ /_/ /___  ____  / /__           \x1b[0m\n", .{});
    std.debug.print("  \x1b[36m / / / / / / / __/ / __ \\/ __ \\/ //_/           \x1b[0m\n", .{});
    std.debug.print("  \x1b[36m/ /_/ / /_/ / /_/ / /_/ / /_/ / ,<              \x1b[0m\n", .{});
    std.debug.print("  \x1b[36m\\____/\\__,_/\\__/_/\\____/\\____/_/|_|             \x1b[0m\n", .{});
    std.debug.print("  \x1b[36m   ____                         _               \x1b[0m\n", .{});
    std.debug.print("  \x1b[36m  / __ \\_________ _____ _____  (_)___ ___  _____\x1b[0m\n", .{});
    std.debug.print("  \x1b[36m / / / / ___/ __ `/ __ `/ __ \\/ /_  // _ \\/ ___/\x1b[0m\n", .{});
    std.debug.print("  \x1b[36m/ /_/ / /  / /_/ / /_/ / / / / / / //  __/ /    \x1b[0m\n", .{});
    std.debug.print("  \x1b[36m\\____/_/   \\__, /\\__,_/_/ /_/_/ /___|\\___/_/     \x1b[0m\n", .{});
    std.debug.print("  \x1b[36m          /____/                                 \x1b[0m\n", .{});
    std.debug.print("\n", .{});
}

pub fn printMenuOptions() void {
    std.debug.print("  \x1b[1;37m1.\x1b[0m Importar PST\n", .{});
    std.debug.print("  \x1b[1;37m2.\x1b[0m Scan PST\n", .{});
    std.debug.print("  \x1b[1;37m0.\x1b[0m Salir\n", .{});
    std.debug.print("\n", .{});
    std.debug.print("  \x1b[33mSelecciona una opcion:\x1b[0m ", .{});
}

pub fn readMenuChoice() !MenuChoice {
    var cursor: usize = 0;
    const labels = [_][]const u8{
        "Importar PST",
        "Scan PST",
        "Salir",
    };

    while (true) {
        clearScreen();
        printBanner();
        std.debug.print("  \x1b[90m↑/↓ mover | Enter confirmar | Q salir\x1b[0m\n\n", .{});

        for (labels, 0..) |label, idx| {
            const is_current = idx == cursor;
            if (is_current) std.debug.print("  \x1b[7m", .{});
            std.debug.print("  {s}\n", .{label});
            if (is_current) std.debug.print("  \x1b[0m", .{});
        }

        const key = readSingleKey() catch continue;
        switch (key) {
            '1' => return .import_pst,
            '2' => return .scan_pst,
            '0', 'q', 'Q' => return .exit,
            'w', 'W', 'k', 'K' => {
                if (cursor > 0) cursor -= 1;
            },
            's', 'S', 'j', 'J' => {
                if (cursor + 1 < labels.len) cursor += 1;
            },
            '\r', '\n' => {
                return switch (cursor) {
                    0 => .import_pst,
                    1 => .scan_pst,
                    else => .exit,
                };
            },
            27 => {
                const seq1 = readSingleKey() catch continue;
                if (seq1 != '[' and seq1 != 'O') continue;

                const seq2 = readSingleKey() catch continue;
                switch (seq2) {
                    'A' => {
                        if (cursor > 0) cursor -= 1;
                    },
                    'B' => {
                        if (cursor + 1 < labels.len) cursor += 1;
                    },
                    else => {},
                }
            },
            0, 224 => {
                const ext = readSingleKey() catch continue;
                switch (ext) {
                    72 => {
                        if (cursor > 0) cursor -= 1;
                    },
                    80 => {
                        if (cursor + 1 < labels.len) cursor += 1;
                    },
                    else => {},
                }
            },
            else => {},
        }
    }
}

pub fn readLine(allocator: std.mem.Allocator) ![]u8 {
    var buf: [1024]u8 = undefined;
    const line = readStdinLine(&buf) catch return error.InvalidInput;
    if (line == null) return error.EndOfStream;

    const trimmed = std.mem.trim(u8, line.?, &[_]u8{ ' ', '\t', '\r', '\n' });
    const result = try allocator.alloc(u8, trimmed.len);
    @memcpy(result, trimmed);
    return result;
}

pub fn readSingleKey() !u8 {
    if (key_pending_pos < key_pending_len) {
        const b = key_pending[key_pending_pos];
        key_pending_pos += 1;
        if (key_pending_pos >= key_pending_len) {
            key_pending_pos = 0;
            key_pending_len = 0;
        }
        return b;
    }

    if (builtin.os.tag == .windows) {
        const input_handle = windows.GetStdHandle(windows.STD_INPUT_HANDLE) catch return error.ReadFailed;

        var original_mode: windows.DWORD = undefined;
        if (windows.kernel32.GetConsoleMode(input_handle, &original_mode) == 0) {
            return error.ReadFailed;
        }

        const ENABLE_PROCESSED_INPUT: windows.DWORD = 0x0001;
        const ENABLE_LINE_INPUT: windows.DWORD = 0x0002;
        const ENABLE_ECHO_INPUT: windows.DWORD = 0x0004;
        const ENABLE_VIRTUAL_TERMINAL_INPUT: windows.DWORD = 0x0200;
        const raw_mode = (original_mode | ENABLE_VIRTUAL_TERMINAL_INPUT) & ~(ENABLE_PROCESSED_INPUT | ENABLE_LINE_INPUT | ENABLE_ECHO_INPUT);
        _ = windows.kernel32.SetConsoleMode(input_handle, raw_mode);
        defer _ = windows.kernel32.SetConsoleMode(input_handle, original_mode);

        const stdin_file = std.fs.File.stdin();
        var many: [16]u8 = undefined;
        while (true) {
            const n = stdin_file.read(&many) catch return error.ReadFailed;
            if (n == 0) return error.EndOfStream;
            if (n > 1) {
                const pending = many[1..n];
                @memcpy(key_pending[0..pending.len], pending);
                key_pending_pos = 0;
                key_pending_len = pending.len;
            }
            return many[0];
        }
    }

    var buf: [8]u8 = undefined;
    const line = readStdinLine(&buf) catch return error.ReadFailed;
    if (line == null or line.?.len == 0) return '\n';
    return line.?[0];
}

pub fn readYesNo(default: bool) bool {
    var buf: [64]u8 = undefined;
    const line = readStdinLine(&buf) catch return default;
    if (line == null) return default;

    const trimmed = std.mem.trim(u8, line.?, &[_]u8{ ' ', '\t', '\r', '\n' });

    if (trimmed.len == 0) return default;

    if (trimmed.len == 1) {
        const c = std.ascii.toLower(trimmed[0]);
        if (c == 's' or c == 'y') return true;
        if (c == 'n') return false;
    }

    // Check "si" or "yes"
    var lower_buf: [16]u8 = undefined;
    const lower_len = @min(trimmed.len, lower_buf.len);
    for (0..lower_len) |i| {
        lower_buf[i] = std.ascii.toLower(trimmed[i]);
    }
    const lower = lower_buf[0..lower_len];

    if (std.mem.eql(u8, lower, "si") or std.mem.eql(u8, lower, "yes")) return true;
    if (std.mem.eql(u8, lower, "no")) return false;

    return default;
}

pub fn printError(msg: []const u8) void {
    std.debug.print("\n  \x1b[1;31mError:\x1b[0m {s}\n", .{msg});
}

pub fn printSuccess(msg: []const u8) void {
    std.debug.print("\n  \x1b[1;32m{s}\x1b[0m\n", .{msg});
}

pub fn printInfo(msg: []const u8) void {
    std.debug.print("  \x1b[36m{s}\x1b[0m\n", .{msg});
}

pub fn printSectionTitle(title: []const u8) void {
    std.debug.print("\n  \x1b[1;33m--- {s} ---\x1b[0m\n\n", .{title});
}

pub fn waitForEnter() void {
    std.debug.print("\n  \x1b[90mPresiona Enter para continuar...\x1b[0m", .{});
    var buf: [64]u8 = undefined;
    _ = readStdinLine(&buf) catch {};
}

pub fn failAbort(msg: []const u8) void {
    printError(msg);
    waitForEnter();
}

pub fn cancelAbort() void {
    std.debug.print("\n  \x1b[90mOperacion cancelada.\x1b[0m\n", .{});
    waitForEnter();
}

fn readStdinLine(buf: []u8) !?[]u8 {
    const stdin_file = std.fs.File.stdin();
    var pos: usize = 0;
    while (pos < buf.len) {
        var one: [1]u8 = undefined;
        const n = stdin_file.read(&one) catch return error.ReadFailed;
        if (n == 0) {
            if (pos == 0) return null;
            return buf[0..pos];
        }
        if (one[0] == '\n') {
            return buf[0..pos];
        }
        buf[pos] = one[0];
        pos += 1;
    }
    return buf[0..pos];
}

pub const DurationParts = struct {
    hours: u64,
    minutes: u64,
    seconds: u64,
};

pub fn secondsToHms(total_seconds: i64) DurationParts {
    const safe_seconds: u64 = @intCast(if (total_seconds < 0) 0 else total_seconds);
    const hours = safe_seconds / 3600;
    const remaining_after_hours = safe_seconds % 3600;
    const minutes = remaining_after_hours / 60;
    const seconds = remaining_after_hours % 60;
    return .{ .hours = hours, .minutes = minutes, .seconds = seconds };
}

pub fn terminalWidthColumns() usize {
    if (builtin.os.tag == .windows) {
        const stdout_handle = windows.GetStdHandle(windows.STD_OUTPUT_HANDLE) catch return 120;

        var csbi: windows.CONSOLE_SCREEN_BUFFER_INFO = undefined;
        if (windows.kernel32.GetConsoleScreenBufferInfo(stdout_handle, &csbi) != 0) {
            const width_i32 = @as(i32, csbi.srWindow.Right) - @as(i32, csbi.srWindow.Left) + 1;
            if (width_i32 > 0) return @intCast(width_i32);
        }
    }

    return 120;
}

pub fn truncateWithEllipsis(text: []const u8, out_buf: []u8, max_visible_chars: usize) []const u8 {
    if (max_visible_chars == 0) return "";
    if (text.len <= max_visible_chars) return text;

    if (max_visible_chars <= 3) {
        const keep = @min(max_visible_chars, out_buf.len);
        if (keep == 0) return "";
        @memcpy(out_buf[0..keep], text[0..keep]);
        return out_buf[0..keep];
    }

    const head = @min(max_visible_chars - 3, out_buf.len - 3);
    if (head == 0) return "";

    @memcpy(out_buf[0..head], text[0..head]);
    out_buf[head + 0] = '.';
    out_buf[head + 1] = '.';
    out_buf[head + 2] = '.';
    return out_buf[0 .. head + 3];
}

pub fn printProgressBar(percent: u32, label: []const u8) void {
    const columns = terminalWidthColumns();
    const safe_columns = if (columns < 20) 20 else columns;

    const fixed_visible = 2 + label.len + 2 + 6;
    const available_for_bar = if (safe_columns > fixed_visible)
        safe_columns - fixed_visible
    else
        8;
    const bar_width: u32 = @intCast(@max(@as(usize, 8), @min(available_for_bar, 40)));

    const filled = (percent * bar_width) / 100;
    const empty = bar_width - filled;

    // Clear line and print label + bar
    std.debug.print("\r\x1b[2K  \x1b[36m{s}\x1b[0m [", .{label});

    var i: u32 = 0;
    while (i < filled) : (i += 1) {
        std.debug.print("\x1b[32m█\x1b[0m", .{});
    }
    i = 0;
    while (i < empty) : (i += 1) {
        std.debug.print("\x1b[90m░\x1b[0m", .{});
    }

    std.debug.print("] {d}%", .{percent});
}
