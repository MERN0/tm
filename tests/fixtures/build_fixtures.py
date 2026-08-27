"""Generates small synthetic xlsx fixtures that mirror the sheet layouts shown
in the user-provided screenshots (feature 019 "Slope Assist" / numeric id #3),
for use in tests and manual CLI runs. These are NOT the real project
workbooks -- they're hand-built approximations reconstructed from photos of a
filtered example, used to validate the extraction logic end-to-end until real
files are available.

Run directly to (re)write the fixture files:
    python tests/fixtures/build_fixtures.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

FIXTURES_DIR = Path(__file__).parent


def _write_sheet(wb: Workbook, name: str, headers: list[str], rows: list[list]) -> None:
    ws = wb.create_sheet(name)
    ws.append(headers)
    for row in rows:
        ws.append(row)


def build_system_requirements() -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    _write_sheet(wb, "Cover Page", ["TE_TMHC_HILS Development & Testing_ System Requirements"], [])

    _write_sheet(
        wb,
        "Index",
        ["Sl No", "Requirement Sheet", "Feature Name", "Function Group"],
        [
            [1, "018", "Low Speed Setting", "Operation Support Function"],
            [2, "019", "Slope Assist", "Operation Support Function"],
        ],
    )

    req_headers = [
        "Requirement ID",
        "Object Heading",
        "Object Text",
        "Category",
        "Variant",
        "Priority",
        "Verification Method",
        "Verification Stage",
        "Verification Criteria",
        "Source",
    ]
    _write_sheet(
        wb,
        "019",
        req_headers,
        [
            ["TMHC_SYSRS_GR019001", "Feature Group", "Slope Assist", "Heading", "-", "-", "-", "-", "-", "-"],
            ["TMHC_SYSRS_GR019002", "Feature Name", "019 - Slope Assist", "Heading", "-", "-", "-", "-", "-", "-"],
            ["TMHC_SYSRS_GR019003", "Feature Details", None, "Heading", "-", "-", "-", "-", "-", "-"],
            [
                "TMHC_SYSRS_GR019004",
                "Purpose",
                "Improve work efficiency and driveability by automatically switching to the "
                "performance equivalent to high power mode when a slope is detected.",
                "Information",
                "-",
                "-",
                "-",
                "-",
                "-",
                "geneB_FunctionSpec (Running Control)",
            ],
            [
                "TMHC_SYSRS_GR019005",
                "Functional Requirement 1",
                "The system shall detect a slope using the Slope_Sensor input when "
                "Slope_Detection_Latency has elapsed.",
                "Functional Requirement",
                "ALL",
                "High",
                "Test",
                "HIL",
                "Slope detected within tolerance",
                "TE_TMHC_HILS_Development_Feature_List_check_20260101",
            ],
            [
                "TMHC_SYSRS_GR019006",
                "Functional Requirement 2",
                "The system shall enable high power mode automatically when a slope is "
                "detected and Slope_Assist_Enabled_Disabled is enabled.",
                "Functional Requirement",
                "ALL",
                "High",
                "Test",
                "HIL",
                "High power mode active",
                "TE_TMHC_HILS_Development_Feature_List_check_20260101",
            ],
        ],
    )

    _write_sheet(
        wb,
        "Master List - Abbreviations",
        ["Abbreviations", "Description/Definition"],
        [
            ["MB Contactor", "Main Battery Contactor"],
            ["RPM", "Rotations per minute"],
            ["VBBT", "Actual Vehicle Battery Voltage"],
            ["VBKY", "Vehicle Battery - Keyed Voltage, battery voltage after the key switch"],
            ["PKB", "Parking Brake"],
            ["MFD", "Multi-Function Display"],
            ["DMC", "Drive Motor Controller"],
            ["PMC", "Pump Motor Controller"],
            ["SOC", "State of charge"],
            ["BMS", "Battery Management System"],
            ["PWM", "Pulse Width Modulated"],
            ["FWD", "Forward"],
            ["BWD", "Backward"],
            ["STD", "Standard"],
            ["HIL", "Hardware in Loop"],
            ["ACCEL", "Accelerator"],
            ["BRK", "Brake"],
            ["OPTSET", "Option Set"],
            ["SOL", "Solenoid"],
        ],
    )

    comm_headers = [
        "Signal ID",
        "Message Name",
        "Message IDs",
        "Logical Signal Name",
        "Signal Name",
        "Signal Description",
        "017",
        "018",
        "019",
        "020",
    ]
    _write_sheet(
        wb,
        "Master Comm Matrix (CAN)",
        comm_headers,
        [
            [
                "TMHC_SYSRS_DBC0001",
                "DrvL_PDO1_Rx",
                "0x211",
                "Commanded_RPM",
                "DrvL_Rx1_CmdSpd",
                "Command motor speed [rpm]",
                "x",
                "x",
                "O",
                "x",
            ],
            [
                "TMHC_SYSRS_DBC0046",
                "Main_SDO_Tx",
                "0x581",
                "Slope_Assist_Enabled_Disabled",
                "Main_TxS_0x2040_0x05",
                "ComData_FunctionFlg4",
                "x",
                "x",
                "O",
                "x",
            ],
            [
                "TMHC_SYSRS_DBC0012",
                "Main_SDO_Tx",
                "0x581",
                "PwrCtrlMode_Tx",
                "Main_TxS_0x2020_0x01",
                "ComData_PwrCtrlMode",
                "x",
                "O",
                "x",
                "x",
            ],
        ],
    )

    param_headers = [
        "Parameter ID",
        "Parameter Name",
        "Parameter Description(purpose)",
        "Parameter Type",
        "Unit",
        "Parameter Value",
        "Parameter default value",
        "017",
        "018",
        "019",
        "020",
    ]
    _write_sheet(
        wb,
        "Master List - App Parameter",
        param_headers,
        [
            [
                "TMHC_SYSRS_PARM0026",
                "Slope_Detection_Latency",
                "Time within which a slope should be detected",
                "Configuration",
                "ms",
                300,
                300,
                "x",
                "x",
                "O",
                "x",
            ],
            [
                "TMHC_SYSRS_PARM0010",
                "Low_Speed_Limit",
                "Speed limit for low speed setting",
                "Configuration",
                "km/h",
                5,
                5,
                "x",
                "O",
                "x",
                "x",
            ],
        ],
    )

    io_headers = ["Signal ID", "Logical Signal Name", "Signal Type", "017", "018", "019", "020"]
    _write_sheet(
        wb,
        "Master Input Output Signals",
        io_headers,
        [
            ["TMHC_SYSRS_IO0005", "Accelerator_Sensor", "Sensor", "x", "x", "O", "x"],
            ["TMHC_SYSRS_IO0009", "Tire_Angle_Sensor", "Sensor", "x", "x", "O", "x"],
            ["TMHC_SYSRS_IO0012", "Power_Select", "CAN", "x", "x", "O", "x"],
            ["TMHC_SYSRS_IO0043", "Slope_Sensor", "Sensor", "x", "x", "O", "x"],
            ["TMHC_SYSRS_IO0002", "Low_Speed_Switch", "Sensor", "x", "O", "x", "x"],
        ],
    )

    path = FIXTURES_DIR / "System Requirements.xlsx"
    wb.save(path)
    return path


def build_command_list() -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    _write_sheet(wb, "Cover Page", ["TE_TMHC_Command_List"], [])

    command_headers = [
        "SL No",
        "Type",
        "Command name",
        "Message Name",
        "Signal Description",
        "Signal Name",
        "Index (Hex)",
        "Subindex (Hex)",
        "Decimal",
        "Platform",
    ]
    _write_sheet(
        wb,
        "Command List",
        command_headers,
        [
            [700, "CAN", "CAN_Main_Deceleration_Command", "Pump_PDO1_Rx", None, "Pump_Rx1_CmdDec", None, None, None, "Platform::Model.Row.MDL..."],
            [706, "CAN", "CAN_Main_D_Accel_Divergence", "Main_SDO_Tx", None, "Main_TxS_0x2920_0x0c", "0x2920", "0x0c", 7365660, "Platform::Model.Row.MDL..."],
            [
                761,
                "CAN",
                "CAN_Main_Slope_AD",
                "Main_SDO_Tx",
                "Slope angle divergence",
                "Main_TxS_0x2470_0x01",
                "0x2470",
                "0x01",
                9344,
                "Platform::Model.Row.MDL...",
            ],
            [
                732,
                "CAN",
                "CAN_HIL_Error_ToDispDiag_2",
                "Main_SDO_Rx",
                "Diagnostic error flag 2",
                "Main_RxS_0x5110_0x03",
                "0x5110",
                "0x03",
                20752,
                "Platform::Model.Row.MDL...",
            ],
            [
                749,
                "CAN",
                "CAN_HIL_Slope_AD",
                "Main_SDO_Rx",
                "Slope angle divergence (HIL)",
                "Main_RxS_0x2470_0x01",
                "0x2470",
                "0x01",
                9344,
                "Platform::Model.Row.MDL...",
            ],
            [
                751,
                "CAN",
                "CAN_HIL_Target_Sound",
                "Disp_PDO2_Rx",
                "Notification tone (2 intermittent)",
                "Main_RxS_SFS_Sound",
                None,
                None,
                None,
                "Platform::Model.Row.MDL...",
            ],
            # Intentionally NOT a command referenced by any Feature sheet row, and
            # CAN_Main_UnknownCommand (referenced below) is intentionally absent from
            # this sheet, to exercise the "command not found" warning path.
        ],
    )

    _write_sheet(
        wb,
        "Feature_ID_Mapping",
        ["Sl No", "Feature Name", "ID"],
        [
            [1, "Maximum Speed Limitation", "#1"],
            [2, "Low Speed Setting", "#2"],
            [3, "Slope Assist", "#3"],
            [4, "BODW MH Prohibition", "#4"],
        ],
    )

    feature_signal_headers = ["SL No", "Type", "Command name", "Signal Name", "#1", "#2", "#3"]
    _write_sheet(
        wb,
        "Recorder_Signals_Feature",
        feature_signal_headers,
        [
            [11, "CAN", "CAN_Main_Deceleration_Command", "Pump_Rx1_CmdDec", True, True, True],
            [14, "CAN", "CAN_Main_D_Accel_Divergence", "Main_TxS_0x2920_0x0c", True, True, True],
            [81, "CAN", "CAN_Main_Slope_AD", "Main_TxS_0x2470_0x01", False, False, True],
            [17, "CAN", "CAN_Main_PwrCtrlMode", "Main_TxS_0x2020_0x01", True, True, False],
            [90, "CAN", "CAN_Main_UnknownCommand", "Main_Test_Unknown", False, False, True],
        ],
    )

    _write_sheet(
        wb,
        "Recorder_Signals_Diag",
        ["SL No", "Type", "Command name", "Signal Name"],
        [[1, "CAN", "CAN_HIL_Error_ToDispDiag_1", "Main_RxS_0x5110_0x02"]],
    )

    _write_sheet(
        wb,
        "SDO_Signals_Feature",
        feature_signal_headers,
        [
            [32, "CAN", "CAN_HIL_Error_ToDispDiag_2", "Main_RxS_0x5110_0x03", False, False, True],
            [44, "CAN", "CAN_HIL_BATT_B48V_B80V", "Main_RxS_0x2464_0x04", False, True, False],
            [49, "CAN", "CAN_HIL_Slope_AD", "Main_RxS_0x2470_0x01", False, False, True],
            [51, "CAN", "CAN_HIL_Target_Sound", "Main_RxS_SFS_Sound", False, False, True],
        ],
    )

    _write_sheet(
        wb,
        "SDO_Signals_Diag",
        ["SL No", "Type", "Command name", "Signal Name"],
        [[1, "CAN", "CAN_HIL_Error_ToDispDiag_1", "Main_RxS_0x5110_0x02"]],
    )

    path = FIXTURES_DIR / "TE_TMHC_Command_List.xlsx"
    wb.save(path)
    return path


def build_configuration_file() -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    _write_sheet(wb, "Cover Page", ["TE_TMHC_Configuration_File"], [])
    _write_sheet(wb, "Tolerances", ["Signal", "Tolerance"], [["Pump_Rx1_CmdDec", "+/- 1%"]])

    model_headers = ["SI No", "Signal", "Test Case Input", "Model Input", "Model Output to ECU", "Remark"]
    ws = wb.create_sheet("Model_Input_Mapping")
    ws.append(model_headers)
    rows = [
        [1, "Pump_Rx1_CmdDec", "ON", 1, "0 (low)", "Deceleration engaged"],
        [2, "Pump_Rx1_CmdDec", "OFF", 0, "0 (low)", "Deceleration released"],
        [3, "Main_TxS_0x2920_0x0c", 30, 30, "2.35 - PS1 (VRA1)", None],
        [4, "Main_TxS_0x2470_0x01", 20, 1, "0V (DSF) - Low", "Slope AD signal"],
        [5, "Main_RxS_0x5110_0x03", "-", "-", "-", "Diagnostic error flag"],
        # Main_RxS_SFS_Sound and Main_RxS_0x2470_0x01 intentionally have no row
        # here, to exercise the "no Model_Input_Mapping match" warning path.
    ]
    for row in rows:
        ws.append(row)
    # Merge the "Signal" column across the two Pump_Rx1_CmdDec rows, mirroring
    # how the real workbook groups repeated signals across several rows.
    ws.merge_cells(start_row=2, start_column=2, end_row=3, end_column=2)

    path = FIXTURES_DIR / "TE_TMHC_Configuration_File.xlsx"
    wb.save(path)
    return path


def build_all() -> list[Path]:
    return [build_system_requirements(), build_command_list(), build_configuration_file()]


if __name__ == "__main__":
    for path in build_all():
        print(f"Wrote {path}")
