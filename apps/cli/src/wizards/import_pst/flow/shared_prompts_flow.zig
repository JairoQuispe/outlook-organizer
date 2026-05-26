const std = @import("std");
const ui = @import("../../../ui.zig");
const types = @import("../types.zig");
const prompts = @import("../prompts.zig");
const routing_wizard = @import("../routing_wizard.zig");
const shared_config_mod = @import("shared_config.zig");

const SharedConfig = shared_config_mod.SharedConfig;

fn promptRoutingOptions(config: *SharedConfig, allocator: std.mem.Allocator) !void {
    ui.clearScreen();
    ui.printSectionTitle("Enrutamiento de Correos");
    std.debug.print("  \x1b[90mDeseas enrutar los correos hacia multiples buzones segun sus fechas?\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mActivar Enrutamiento de Correos? (s/N) [N]:\x1b[0m ", .{});
    config.enable_routing = ui.readYesNo(false);

    if (config.enable_routing) {
        config.routing_criterion = routing_wizard.selectRoutingCriterion() catch |err| {
            if (err == error.Cancelled) return error.Cancelled;
            ui.failAbort("Error seleccionando criterio de enrutamiento");
            return error.OperationAborted;
        };
    }

    config.scan_mode = if (config.enable_routing) .deep else (prompts.chooseScanMode(allocator) catch {
        ui.failAbort("Error leyendo modo de escaneo");
        return error.OperationAborted;
    });
}

fn promptDuplicateOptions(config: *SharedConfig) void {
    ui.clearScreen();
    ui.printSectionTitle("Duplicados");
    std.debug.print("  \x1b[90mSaltar items duplicados detectados en el buzon destino?\x1b[0m\n", .{});
    std.debug.print("  \x1b[90mUsa Message-ID, SearchKey o clave compuesta para detectar.\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mSaltar duplicados? (S/n) [S]:\x1b[0m ", .{});
    config.skip_duplicates = ui.readYesNo(true);

    if (config.skip_duplicates) {
        std.debug.print("\n  \x1b[90mRevision profunda: tambien indexa subcarpetas del destino.\x1b[0m\n", .{});
        std.debug.print("  \x1b[90m(Mas lento, pero detecta duplicados movidos manualmente)\x1b[0m\n\n", .{});
        std.debug.print("  \x1b[33mRevision profunda? (s/N) [N]:\x1b[0m ", .{});
        config.deep_duplicate_check = ui.readYesNo(false);
    }
}

fn promptFilterAndPerformanceOptions(config: *SharedConfig, allocator: std.mem.Allocator) void {
    config.filter_year = promptOptionalYearFilter(allocator);
    config.filter_months = promptOptionalMonthsFilter(allocator);

    ui.clearScreen();
    ui.printSectionTitle("Rendimiento");
    std.debug.print("  \x1b[90mThrottling adaptativo: ajusta la velocidad automaticamente\x1b[0m\n", .{});
    std.debug.print("  \x1b[90mcuando Exchange limita las peticiones. Recomendado para\x1b[0m\n", .{});
    std.debug.print("  \x1b[90mExchange Online / Office 365.\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mActivar throttling adaptativo? (S/n) [S]:\x1b[0m ", .{});
    config.adaptive_throttling = ui.readYesNo(true);
}

fn promptOptionalYearFilter(allocator: std.mem.Allocator) ?[]const u8 {
    ui.clearScreen();
    ui.printSectionTitle("Filtros de fecha");
    std.debug.print("  \x1b[90mFiltrar por anio? Dejar vacio para importar todos.\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mAnio (ej: 2023) o Enter para todos:\x1b[0m ", .{});

    const year_input = ui.readLine(allocator) catch return null;
    if (year_input.len == 0) {
        allocator.free(year_input);
        return null;
    }

    const year_value = std.fmt.parseInt(i32, year_input, 10) catch {
        ui.printError("Anio invalido, se ignorara el filtro");
        allocator.free(year_input);
        ui.waitForEnter();
        return null;
    };

    if (year_value < 1900 or year_value > 9999) {
        ui.printError("Anio fuera de rango (1900-9999), se ignorara el filtro");
        allocator.free(year_input);
        ui.waitForEnter();
        return null;
    }

    return year_input;
}

fn promptOptionalMonthsFilter(allocator: std.mem.Allocator) ?[]const u8 {
    std.debug.print("\n  \x1b[90mFiltrar por meses? Separar con comas.\x1b[0m\n", .{});
    std.debug.print("  \x1b[90mAcepta: numeros (1-12), nombres (enero, feb, march)\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mMeses (ej: ene,feb,mar) o Enter para todos:\x1b[0m ", .{});

    const months_input = ui.readLine(allocator) catch return null;
    if (months_input.len == 0) {
        allocator.free(months_input);
        return null;
    }

    return months_input;
}

pub fn promptSharedConfig(allocator: std.mem.Allocator, reference_pst_path: []const u8) !SharedConfig {
    var config = SharedConfig.initDefaults();

    config.profile_name = prompts.chooseOutlookProfile(allocator) catch {
        ui.failAbort("Error seleccionando perfil de Outlook");
        return error.OperationAborted;
    };

    promptRoutingOptions(&config, allocator) catch |err| {
        if (err == error.Cancelled) return error.Cancelled;
        return err;
    };

    config.scan_filter_year = prompts.chooseScanYearFilter(allocator) catch {
        ui.failAbort("Error leyendo filtro de anio para escaneo");
        return error.OperationAborted;
    };

    config.action = prompts.chooseImportAction(reference_pst_path);

    promptDuplicateOptions(&config);
    promptFilterAndPerformanceOptions(&config, allocator);

    return config;
}
