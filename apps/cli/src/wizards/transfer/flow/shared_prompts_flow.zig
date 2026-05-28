const std = @import("std");
const ui = @import("../../../ui.zig");
const types = @import("../types.zig");
const prompts = @import("../../import_pst/prompts.zig");
const routing_wizard = @import("../../import_pst/routing_wizard.zig");
const shared_config_mod = @import("shared_config.zig");

pub const SharedConfig = shared_config_mod.SharedConfig;

pub fn promptSharedConfig(allocator: std.mem.Allocator) !SharedConfig {
    var config = SharedConfig.initDefaults();

    config.profile_name = prompts.chooseOutlookProfile(allocator) catch {
        ui.failAbort("Error seleccionando perfil de Outlook");
        return error.OperationAborted;
    };

    // Prompt for routing options
    ui.clearScreen();
    ui.printSectionTitle("Enrutamiento de Correos");
    std.debug.print("  \x1b[90mDeseas enrutar los correos hacia multiples buzones segun sus fechas?\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mActivar Enrutamiento de Correos? (s/N) [N]:\x1b[0m ", .{});
    config.enable_routing = ui.readYesNo(false);

    if (config.enable_routing) {
        const criterion = routing_wizard.selectRoutingCriterion() catch |err| {
            if (err == error.Cancelled) return error.Cancelled;
            ui.failAbort("Error seleccionando criterio de enrutamiento");
            return error.OperationAborted;
        };
        config.routing_criterion = @as(types.RoutingCriterion, @enumFromInt(@intFromEnum(criterion)));
    }

    config.scan_mode = if (config.enable_routing) .deep else (block: {
        const mode = prompts.chooseScanMode(allocator) catch {
            ui.failAbort("Error leyendo modo de escaneo");
            return error.OperationAborted;
        };
        break :block @as(types.ScanMode, @enumFromInt(@intFromEnum(mode)));
    });

    config.scan_filter_year = prompts.chooseScanYearFilter(allocator) catch {
        ui.failAbort("Error leyendo filtro de anio para escaneo");
        return error.OperationAborted;
    };

    config.action = promptAction();

    ui.clearScreen();
    ui.printSectionTitle("Duplicados");
    std.debug.print("  \x1b[90mSaltar items duplicados detectados en el destino?\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mSaltar duplicados? (S/n) [S]:\x1b[0m ", .{});
    config.skip_duplicates = ui.readYesNo(true);

    if (config.skip_duplicates) {
        std.debug.print("\n  \x1b[90mRevision profunda: tambien indexa subcarpetas del destino.\x1b[0m\n\n", .{});
        std.debug.print("  \x1b[33mRevision profunda? (s/N) [N]:\x1b[0m ", .{});
        config.deep_duplicate_check = ui.readYesNo(false);
    }

    ui.clearScreen();
    ui.printSectionTitle("Filtros de fecha");
    std.debug.print("  \x1b[90mFiltrar por anio? Dejar vacio para transferir todos.\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mAnio (ej: 2023) o Enter para todos:\x1b[0m ", .{});
    {
        const year_input = ui.readLine(allocator) catch null;
        if (year_input) |y| {
            if (y.len > 0) {
                const val = std.fmt.parseInt(i32, y, 10) catch {
                    allocator.free(y);
                    return error.InvalidYear;
                };
                if (val >= 1900 and val <= 9999) {
                    config.filter_year = y;
                } else {
                    allocator.free(y);
                }
            } else {
                allocator.free(y);
            }
        }
    }

    std.debug.print("\n  \x1b[90mFiltrar por meses? Separar con comas.\x1b[0m\n", .{});
    std.debug.print("  \x1b[33mMeses (ej: ene,feb,mar) o Enter para todos:\x1b[0m ", .{});
    {
        const months_input = ui.readLine(allocator) catch null;
        if (months_input) |m| {
            if (m.len > 0) {
                config.filter_months = m;
            } else {
                allocator.free(m);
            }
        }
    }

    ui.clearScreen();
    ui.printSectionTitle("Rendimiento");
    std.debug.print("  \x1b[90mThrottling adaptativo: ajusta la velocidad automaticamente\x1b[0m\n", .{});
    std.debug.print("  \x1b[90mcuando Exchange limita las peticiones.\x1b[0m\n\n", .{});
    std.debug.print("  \x1b[33mActivar throttling adaptativo? (S/n) [S]:\x1b[0m ", .{});
    config.adaptive_throttling = ui.readYesNo(true);

    return config;
}

pub fn promptAction() []const u8 {
    var cursor: usize = 0;
    const labels = [_][]const u8{
        "Copiar (los originales permanecen en el origen)",
        "Mover (los originales se eliminan del origen)",
    };

    while (true) {
        ui.clearScreen();
        ui.printSectionTitle("Accion");
        std.debug.print("  \x1b[90m↑/↓ mover | Enter confirmar\x1b[0m\n\n", .{});

        for (labels, 0..) |label, idx| {
            const is_current = idx == cursor;
            if (is_current) std.debug.print("  \x1b[7m", .{});
            std.debug.print("  {s}", .{label});
            if (is_current) std.debug.print("\x1b[0m", .{});
            std.debug.print("\n", .{});
        }

        const input = ui.readMenuInput(&cursor, labels.len) catch continue;
        switch (input) {
            .cancel => return "Copy",
            .enter => return if (cursor == 0) "Copy" else "Move",
            else => {},
        }
    }
}
