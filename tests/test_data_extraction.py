from __future__ import annotations

from pathlib import Path

from tests.fixtures.build_fixtures import build_all

from src.graph.build_graph import compiled_graph

FIXTURES_DIR = Path(__file__).parent / "fixtures"


def _run(feature_id: str) -> dict:
    build_all()  # ensure fixtures exist / are up to date
    state = {
        "feature_id": feature_id,
        "sysreq_path": str(FIXTURES_DIR / "System Requirements.xlsx"),
        "command_list_path": str(FIXTURES_DIR / "TE_TMHC_Command_List.xlsx"),
        "config_path": str(FIXTURES_DIR / "TE_TMHC_Configuration_File.xlsx"),
        "output_dir": "",
        "extracted": {},
        "warnings": [],
    }
    return compiled_graph.invoke(state)


def test_feature_name_resolved_from_index():
    result = _run("019")
    assert result["extracted"]["feature_name"] == "Slope Assist"


def test_functional_requirements_filtered():
    result = _run("019")
    reqs = result["extracted"]["requirements"]["functional_requirements"]
    assert len(reqs) == 2
    assert all(r["category"] == "Functional Requirement" for r in reqs)
    ids = {r["requirement id"] for r in reqs}
    assert ids == {"TMHC_SYSRS_GR019005", "TMHC_SYSRS_GR019006"}
    # Heading/Information rows must not leak into functional_requirements
    assert all("Heading" not in r["object heading"] for r in reqs)


def test_requirement_metadata_captured_separately():
    result = _run("019")
    metadata = result["extracted"]["requirements"]["metadata"]
    headings = {m["object heading"] for m in metadata}
    assert {"Feature Group", "Feature Name", "Purpose"} <= headings


def test_comm_matrix_filtered_to_valid_marker_only():
    result = _run("019")
    comm_matrix = result["extracted"]["comm_matrix"]
    names = {row["logical signal name"] for row in comm_matrix}
    assert names == {"Commanded_RPM", "Slope_Assist_Enabled_Disabled"}
    # PwrCtrlMode_Tx is only valid for 018 ('O'), must be excluded for 019
    assert "PwrCtrlMode_Tx" not in names
    # marker columns should be stripped from the output rows
    for row in comm_matrix:
        assert "019" not in row
        assert "018" not in row


def test_app_parameters_filtered():
    result = _run("019")
    params = result["extracted"]["app_parameters"]
    names = {p["parameter name"] for p in params}
    assert names == {"Slope_Detection_Latency"}


def test_io_signals_filtered():
    result = _run("019")
    io_signals = result["extracted"]["io_signals"]
    names = {s["logical signal name"] for s in io_signals}
    assert names == {"Accelerator_Sensor", "Tire_Angle_Sensor", "Power_Select", "Slope_Sensor"}
    assert "Low_Speed_Switch" not in names


def test_numeric_feature_id_resolved():
    result = _run("019")
    assert result["extracted"]["numeric_feature_id"] == "#3"


def test_recorder_and_sdo_signals_checkbox_filtered():
    result = _run("019")
    extracted = result["extracted"]
    recorder_commands = {row["command name"] for row in extracted["recorder_signals"]}
    assert recorder_commands == {
        "CAN_Main_Deceleration_Command",
        "CAN_Main_D_Accel_Divergence",
        "CAN_Main_Slope_AD",
        "CAN_Main_UnknownCommand",
    }
    # CAN_Main_PwrCtrlMode is checked for #1/#2 but not #3 - must be excluded
    assert "CAN_Main_PwrCtrlMode" not in recorder_commands

    sdo_commands = {row["command name"] for row in extracted["sdo_signals"]}
    assert sdo_commands == {
        "CAN_HIL_Error_ToDispDiag_2",
        "CAN_HIL_Slope_AD",
        "CAN_HIL_Target_Sound",
    }
    assert "CAN_HIL_BATT_B48V_B80V" not in sdo_commands


def test_command_details_and_model_input_mapping():
    result = _run("019")
    commands = {c["command_name"]: c for c in result["extracted"]["commands"]}

    matched = commands["CAN_Main_Deceleration_Command"]
    assert matched["command_list_details"]["signal name"] == "Pump_Rx1_CmdDec"
    assert len(matched["model_input_mapping"]) == 2  # merged-cell grouped rows both recovered
    assert {m["test case input"] for m in matched["model_input_mapping"]} == {"ON", "OFF"}

    unmatched_model_input = commands["CAN_HIL_Target_Sound"]
    assert unmatched_model_input["command_list_details"] is not None
    assert unmatched_model_input["model_input_mapping"] == []

    missing_command = commands["CAN_Main_UnknownCommand"]
    assert missing_command["command_list_details"] is None
    assert missing_command["model_input_mapping"] == []


def test_warnings_surface_unmatched_lookups():
    result = _run("019")
    warnings = " | ".join(result["warnings"])
    assert "CAN_Main_UnknownCommand" in warnings
    assert "Main_RxS_SFS_Sound" in warnings


def test_feature_with_no_requirement_sheet_reports_warning_not_crash():
    result = _run("999")
    assert result["extracted"]["requirements"] == {"metadata": [], "functional_requirements": []}
    assert any("999" in w for w in result["warnings"])
