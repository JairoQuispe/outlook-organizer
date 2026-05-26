const std = @import("std");

pub const RoutingCriterion = enum {
    by_year,
    by_month,
};

pub const TargetStoreMapping = struct {
    year: i32,
    month: ?u8, // 1-12, null if Criterion == .by_year
    store_id: []const u8,
    store_name: []const u8,
    store_type: []const u8,

    pub fn deinit(self: TargetStoreMapping, allocator: std.mem.Allocator) void {
        allocator.free(self.store_id);
        allocator.free(self.store_name);
        allocator.free(self.store_type);
    }
};

pub const ScanMode = enum {
    quick,
    deep,
};

pub const ImportProgressState = struct {
    start_ms: i64,
    copied: i64,
    moved: i64,
    skipped: i64,
    failed: i64,
    size_bytes: i64,
    percent: u32,
    has_rendered_progress: bool,
};

pub const ScannedFolder = struct {
    path: []u8,
    item_count: i64,
    size_bytes: ?i64,
    year_breakdown_display: []u8,
};

pub const FolderTreeNode = struct {
    name: []u8,
    full_path: []u8,
    parent: ?usize,
    children: std.ArrayListUnmanaged(usize),
    expanded: bool,
    folder_index: ?usize,
};

pub const VisibleTreeRow = struct {
    node_index: usize,
    depth: usize,
};

pub const FolderTree = struct {
    nodes: std.ArrayListUnmanaged(FolderTreeNode),
    roots: std.ArrayListUnmanaged(usize),
};

pub const FolderSelectionResult = struct {
    folder_plan_path: []u8,
    scan_export_path: []u8,
    selected_count: usize,
    total_count: usize,
    scan_mode: ScanMode,
};

pub const ImportConfig = struct {
    pst_path: []const u8,
    target_store_id: []const u8,
    target_store_name: []const u8,
    target_store_type: []const u8,
    action: []const u8,
    skip_duplicates: bool,
    deep_duplicate_check: bool,
    filter_year: ?[]const u8,
    filter_months: ?[]const u8,
    folder_plan_path: []const u8,
    adaptive_throttling: bool,
    profile_name: ?[]const u8,
    routing_criterion: ?RoutingCriterion,
    routing_mappings: ?[]const TargetStoreMapping,
};
