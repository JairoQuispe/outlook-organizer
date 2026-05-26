const std = @import("std");

pub const ScriptId = enum {
    import_pst,
    list_stores,
    scan_pst,
};

pub const ScriptRunResult = struct {
    output: []u8,
    command_line: []u8,
    exit_code: u32 = 0,
};

pub const LineCallback = *const fn (ctx: *anyopaque, line: []const u8) void;

/// Writes an embedded script to a temp file and returns its path.
/// Caller owns the returned path memory.
pub fn writeEmbeddedScript(allocator: std.mem.Allocator, script_id: ScriptId) ![]u8 {
    const content = switch (script_id) {
        .import_pst => @embedFile("outlook-import-pst.ps1"),
        .list_stores => @embedFile("outlook-list-stores.ps1"),
        .scan_pst => @embedFile("outlook-scan-pst.ps1"),
    };

    const script_tag = switch (script_id) {
        .import_pst => "import-pst",
        .list_stores => "list-stores",
        .scan_pst => "scan-pst",
    };

    const pid: u32 = std.os.windows.GetCurrentProcessId();
    const ts: u64 = @intCast(@max(@as(i64, @intCast(std.time.milliTimestamp())), 0));
    const filename = try std.fmt.allocPrint(allocator, "oo-{s}-{d}-{d}.ps1", .{ script_tag, pid, ts });
    defer allocator.free(filename);

    // Use Windows TEMP directory
    const temp_dir = std.process.getEnvVarOwned(allocator, "TEMP") catch |err| {
        if (err == error.EnvironmentVariableNotFound) {
            return try allocator.dupe(u8, filename);
        }
        return err;
    };
    defer allocator.free(temp_dir);

    const path = try std.fs.path.join(allocator, &.{ temp_dir, filename });

    // Write script content to temp file
    const file = std.fs.createFileAbsolute(path, .{ .truncate = true }) catch {
        allocator.free(path);
        return error.CannotWriteScript;
    };
    defer file.close();

    const utf8_bom = [_]u8{ 0xEF, 0xBB, 0xBF };
    const has_bom = content.len >= 3 and
        content[0] == utf8_bom[0] and
        content[1] == utf8_bom[1] and
        content[2] == utf8_bom[2];

    if (!has_bom) {
        file.writeAll(&utf8_bom) catch {
            allocator.free(path);
            return error.CannotWriteScript;
        };
    }

    file.writeAll(content) catch {
        allocator.free(path);
        return error.CannotWriteScript;
    };

    return path;
}

/// Runs a PowerShell script with arguments and captures stdout.
/// Returns the full stdout output. Caller owns the memory.
pub fn runScript(allocator: std.mem.Allocator, script_path: []const u8, args: []const []const u8) ![]u8 {
    const result = try runScriptDetailed(allocator, script_path, args);
    defer allocator.free(result.command_line);
    return result.output;
}

pub fn runScriptDetailed(allocator: std.mem.Allocator, script_path: []const u8, args: []const []const u8) !ScriptRunResult {
    return runScriptDetailedStreaming(allocator, script_path, args, null, null);
}

pub fn runScriptDetailedStreaming(
    allocator: std.mem.Allocator,
    script_path: []const u8,
    args: []const []const u8,
    callback: ?LineCallback,
    callback_ctx: ?*anyopaque,
) !ScriptRunResult {
    return runScriptWithHostDetailed(allocator, "powershell.exe", script_path, args, callback, callback_ctx) catch |err| switch (err) {
        error.FileNotFound => runScriptWithHostDetailed(allocator, "pwsh", script_path, args, callback, callback_ctx),
        else => err,
    };
}

pub fn buildCommandPreview(
    allocator: std.mem.Allocator,
    host: []const u8,
    script_path: []const u8,
    args: []const []const u8,
) ![]u8 {
    var argv = std.ArrayListUnmanaged([]const u8){};
    defer argv.deinit(allocator);

    try argv.appendSlice(allocator, &.{
        host,
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        script_path,
    });
    try argv.appendSlice(allocator, args);

    return formatCommandLine(allocator, argv.items);
}

fn runScriptWithHostDetailed(
    allocator: std.mem.Allocator,
    host: []const u8,
    script_path: []const u8,
    args: []const []const u8,
    callback: ?LineCallback,
    callback_ctx: ?*anyopaque,
) !ScriptRunResult {
    // Build argv: <host> -ExecutionPolicy Bypass -File <script> <args...>
    var argv = std.ArrayListUnmanaged([]const u8){};
    defer argv.deinit(allocator);

    try argv.appendSlice(allocator, &.{
        host,
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        script_path,
    });

    for (args) |arg| {
        try argv.append(allocator, arg);
    }

    const command_line = try formatCommandLine(allocator, argv.items);
    errdefer allocator.free(command_line);

    var child = std.process.Child.init(argv.items, allocator);
    child.stdout_behavior = .Pipe;
    child.stderr_behavior = .Inherit;

    try child.spawn();

    // Read all stdout
    var stdout_list = std.ArrayListUnmanaged(u8){};
    defer stdout_list.deinit(allocator);
    var pending_line = std.ArrayListUnmanaged(u8){};
    defer pending_line.deinit(allocator);

    if (child.stdout) |stdout_file| {
        var buf: [4096]u8 = undefined;
        while (true) {
            const n = stdout_file.read(&buf) catch break;
            if (n == 0) break;
            const chunk = buf[0..n];
            try stdout_list.appendSlice(allocator, chunk);

            if (callback != null and callback_ctx != null) {
                try pending_line.appendSlice(allocator, chunk);
                while (std.mem.indexOfScalar(u8, pending_line.items, '\n')) |nl_idx| {
                    const raw = pending_line.items[0..nl_idx];
                    const line = std.mem.trimRight(u8, raw, "\r");
                    callback.?(callback_ctx.?, line);

                    const remaining = pending_line.items.len - (nl_idx + 1);
                    if (remaining > 0) {
                        std.mem.copyForwards(u8, pending_line.items[0..remaining], pending_line.items[nl_idx + 1 ..]);
                    }
                    pending_line.items.len = remaining;
                }
            }
        }

        if (callback != null and callback_ctx != null and pending_line.items.len > 0) {
            const line = std.mem.trimRight(u8, pending_line.items, "\r");
            callback.?(callback_ctx.?, line);
        }
    }

    const term = child.wait() catch {
        return error.ProcessFailed;
    };
    var exit_code: u32 = 0;
    switch (term) {
        .Exited => |code| {
            exit_code = code;
        },
        else => return error.ProcessFailed,
    }

    return .{
        .output = try allocator.dupe(u8, stdout_list.items),
        .command_line = command_line,
        .exit_code = exit_code,
    };
}

fn formatCommandLine(allocator: std.mem.Allocator, argv: []const []const u8) ![]u8 {
    var out = std.ArrayListUnmanaged(u8){};
    errdefer out.deinit(allocator);

    for (argv, 0..) |arg, idx| {
        if (idx > 0) try out.append(allocator, ' ');
        try appendShellArg(allocator, &out, arg);
    }

    return try out.toOwnedSlice(allocator);
}

fn appendShellArg(allocator: std.mem.Allocator, out: *std.ArrayListUnmanaged(u8), arg: []const u8) !void {
    const needs_quotes = std.mem.indexOfAny(u8, arg, " \t\"") != null;
    if (!needs_quotes) {
        try out.appendSlice(allocator, arg);
        return;
    }

    try out.append(allocator, '"');
    for (arg) |ch| {
        if (ch == '"') try out.append(allocator, '\\');
        try out.append(allocator, ch);
    }
    try out.append(allocator, '"');
}

/// Cleanup temp scripts
pub fn cleanupScript(allocator: std.mem.Allocator, path: []const u8) void {
    std.fs.deleteFileAbsolute(path) catch {};
    allocator.free(path);
}
