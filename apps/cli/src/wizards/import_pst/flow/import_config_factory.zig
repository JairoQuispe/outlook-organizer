const types = @import("../types.zig");
const shared_config_mod = @import("shared_config.zig");

const SharedConfig = shared_config_mod.SharedConfig;

pub fn buildImportConfig(
    shared: *const SharedConfig,
    pst_path: []const u8,
    target_store_id: []const u8,
    target_store_name: []const u8,
    target_store_type: []const u8,
    folder_plan_path: []const u8,
    routing_mappings: ?[]const types.TargetStoreMapping,
) types.ImportConfig {
    return .{
        .pst_path = pst_path,
        .target_store_id = target_store_id,
        .target_store_name = target_store_name,
        .target_store_type = target_store_type,
        .action = shared.action,
        .skip_duplicates = shared.skip_duplicates,
        .deep_duplicate_check = shared.deep_duplicate_check,
        .filter_year = shared.filter_year,
        .filter_months = shared.filter_months,
        .folder_plan_path = folder_plan_path,
        .adaptive_throttling = shared.adaptive_throttling,
        .profile_name = shared.profile_name,
        .routing_criterion = shared.routing_criterion,
        .routing_mappings = routing_mappings,
    };
}
