const std = @import("std");
const ui = @import("ui.zig");
const ps_runner = @import("ps_runner.zig");

pub const StoreInfo = struct {
    display_name: []const u8,
    store_id: []const u8,
    file_path: []const u8,
    file_size: []const u8,
    store_type: []const u8,
};

pub const SelectedStore = struct {
    display_name: []u8,
    store_id: []u8,
    store_type: []u8,
};

/// Lists Outlook stores and lets the user pick a destination mailbox.
/// Returns selected store id + display name. Caller owns both fields.
pub fn selectTargetStore(allocator: std.mem.Allocator, profile_name: ?[]const u8) !SelectedStore {
    ui.clearScreen();
    ui.printSectionTitle("Buzon de destino");
    std.debug.print("  \x1b[90mObteniendo buzones de Outlook...\x1b[0m\n\n", .{});

    // Write and run the list-stores script
    const script_path = try ps_runner.writeEmbeddedScript(allocator, .list_stores);
    defer ps_runner.cleanupScript(allocator, script_path);

    var args = std.ArrayListUnmanaged([]const u8){};
    defer args.deinit(allocator);

    try args.append(allocator, "-Json");
    if (profile_name) |p| {
        if (p.len > 0) {
            try args.append(allocator, "-ProfileName");
            try args.append(allocator, p);
        }
    }

    const output = ps_runner.runScript(allocator, script_path, args.items) catch {
        ui.printError("No se pudo ejecutar outlook-list-stores.ps1");
        return error.ScriptFailed;
    };
    defer allocator.free(output);

    // Parse JSON output to find stores
    var stores = std.ArrayListUnmanaged(StoreInfo){};
    defer {
        for (stores.items) |store| {
            allocator.free(store.display_name);
            allocator.free(store.store_id);
            allocator.free(store.file_path);
            allocator.free(store.file_size);
            allocator.free(store.store_type);
        }
        stores.deinit(allocator);
    }

    // Parse the JSON - find lines containing "stores" type
    var line_iter = std.mem.splitSequence(u8, output, "\n");
    while (line_iter.next()) |raw_line| {
        const line = std.mem.trim(u8, raw_line, &[_]u8{ ' ', '\t', '\r', '\n' });
        if (line.len == 0) continue;

        // Try to parse as JSON
        if (std.mem.indexOf(u8, line, "\"type\"") == null) continue;
        if (std.mem.indexOf(u8, line, "\"stores\"") == null) continue;

        // Found the stores payload - parse it
        parseStoresJson(allocator, line, &stores) catch continue;
        break;
    }

    if (stores.items.len == 0) {
        ui.printError("No se encontraron buzones en Outlook");
        std.debug.print("\n  \x1b[90mSalida del script:\x1b[0m\n", .{});
        // Show first 500 chars of output for debugging
        const show_len = @min(output.len, 500);
        std.debug.print("  {s}\n", .{output[0..show_len]});
        return error.NoStoresFound;
    }

    var cursor: usize = 0;
    var scroll: usize = 0;
    const page_size: usize = 16;

    while (true) {
        if (cursor < scroll) scroll = cursor;
        if (cursor >= scroll + page_size) scroll = cursor - page_size + 1;

        ui.clearScreen();
        ui.printSectionTitle("Buzon de destino");
        std.debug.print("  \x1b[90m↑/↓ mover | Enter confirmar | Q cancelar\x1b[0m\n\n", .{});

        const start = scroll;
        const end = @min(stores.items.len, scroll + page_size);
        for (stores.items[start..end], start..) |store, idx| {
            const is_current = idx == cursor;
            if (is_current) std.debug.print("  \x1b[7m", .{});

            std.debug.print("  {s}", .{store.display_name});
            if (store.store_type.len > 0) {
                std.debug.print(" \x1b[90m({s})\x1b[0m", .{store.store_type});
            }
            if (store.file_size.len > 0) {
                std.debug.print(" \x1b[90m[{s}]\x1b[0m", .{store.file_size});
            }

            if (is_current) std.debug.print("\x1b[0m", .{});
            std.debug.print("\n", .{});
        }

        if (end < stores.items.len) {
            std.debug.print("\n  \x1b[90m... y {d} mas\x1b[0m\n", .{stores.items.len - end});
        }

        const input = ui.readMenuInput(&cursor, stores.items.len) catch continue;
        switch (input) {
            .cancel => return error.Cancelled,
            .enter => break,
            else => {},
        }
    }

    const selected = stores.items[cursor];
    return .{
        .display_name = try allocator.dupe(u8, selected.display_name),
        .store_id = try allocator.dupe(u8, selected.store_id),
        .store_type = try allocator.dupe(u8, selected.store_type),
    };
}

fn parseStoresJson(allocator: std.mem.Allocator, json_str: []const u8, stores: *std.ArrayListUnmanaged(StoreInfo)) !void {
    // Simple JSON parsing - look for store objects with displayName and storeId
    // Format: {"type":"stores","stores":[{"displayName":"...","storeId":"...","filePath":"...","fileSize":"...","storeType":"..."},...]}

    // Find the stores array
    const stores_key = "\"stores\"";
    const stores_start = std.mem.indexOf(u8, json_str, stores_key) orelse return;
    const after_key = stores_start + stores_key.len;

    // Find the opening bracket
    var pos = after_key;
    while (pos < json_str.len and json_str[pos] != '[' and json_str[pos] != '{') : (pos += 1) {}
    if (pos >= json_str.len) return;

    const is_array = json_str[pos] == '[';
    if (is_array) {
        pos += 1; // skip '['
    }

    // Parse each store object
    while (pos < json_str.len) {
        // Skip whitespace and commas
        while (pos < json_str.len and (json_str[pos] == ' ' or json_str[pos] == ',' or json_str[pos] == '\t' or json_str[pos] == '\n' or json_str[pos] == '\r')) : (pos += 1) {}

        if (pos >= json_str.len) break;
        if (is_array and json_str[pos] == ']') break;

        if (json_str[pos] != '{') break;

        // Find the matching closing brace
        var depth: u32 = 0;
        const obj_start = pos;
        var in_string = false;
        var escape_next = false;
        while (pos < json_str.len) {
            const c = json_str[pos];
            if (escape_next) {
                escape_next = false;
                pos += 1;
                continue;
            }
            if (c == '\\' and in_string) {
                escape_next = true;
                pos += 1;
                continue;
            }
            if (c == '"') {
                in_string = !in_string;
            } else if (!in_string) {
                if (c == '{') {
                    depth += 1;
                } else if (c == '}') {
                    depth -= 1;
                    if (depth == 0) {
                        pos += 1;
                        break;
                    }
                }
            }
            pos += 1;
        }

        const obj_str = json_str[obj_start..pos];

        const display_name = extractJsonString(allocator, obj_str, "displayName") catch try allocator.dupe(u8, "");
        const store_id = extractJsonString(allocator, obj_str, "storeId") catch try allocator.dupe(u8, "");
        const file_path = extractJsonString(allocator, obj_str, "filePath") catch try allocator.dupe(u8, "");
        const file_size = extractJsonString(allocator, obj_str, "fileSize") catch try allocator.dupe(u8, "");
        const store_type = extractJsonString(allocator, obj_str, "storeType") catch try allocator.dupe(u8, "");

        if (store_id.len > 0) {
            try stores.append(allocator, .{
                .display_name = display_name,
                .store_id = store_id,
                .file_path = file_path,
                .file_size = file_size,
                .store_type = store_type,
            });
        } else {
            allocator.free(display_name);
            allocator.free(store_id);
            allocator.free(file_path);
            allocator.free(file_size);
            allocator.free(store_type);
        }

        if (!is_array) {
            break;
        }
    }
}

fn extractJsonString(allocator: std.mem.Allocator, json: []const u8, key: []const u8) ![]u8 {
    // Find "key":"value" pattern
    // Build search pattern: "key"
    var search_buf: [256]u8 = undefined;
    const search = std.fmt.bufPrint(&search_buf, "\"{s}\"", .{key}) catch return error.KeyTooLong;

    const key_pos = std.mem.indexOf(u8, json, search) orelse return error.KeyNotFound;
    var pos = key_pos + search.len;

    // Skip : and whitespace
    while (pos < json.len and (json[pos] == ':' or json[pos] == ' ' or json[pos] == '\t')) : (pos += 1) {}

    if (pos >= json.len) return error.ValueNotFound;

    // Handle null
    if (pos + 4 <= json.len and std.mem.eql(u8, json[pos .. pos + 4], "null")) {
        return try allocator.dupe(u8, "");
    }

    // Expect opening quote
    if (json[pos] != '"') return error.ValueNotFound;
    pos += 1;

    // Find closing quote (handling escapes)
    const value_start = pos;
    var escape = false;
    while (pos < json.len) {
        if (escape) {
            escape = false;
            pos += 1;
            continue;
        }
        if (json[pos] == '\\') {
            escape = true;
            pos += 1;
            continue;
        }
        if (json[pos] == '"') break;
        pos += 1;
    }

    return try allocator.dupe(u8, json[value_start..pos]);
}
