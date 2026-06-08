"""
Simplified AI engine – adds dummy sheets for A1-A10 (works without numpy/pandas)
"""
from openpyxl import Workbook

def add_ai_sheets_to_workbook(wb, coils_df, inspections_df, inspectors_df, line_stats, avail_insp=3):
    """
    Adds sheets A1 to A10 to an openpyxl Workbook (simple dummy data).
    """
    # A1 Demand Forecast
    ws = wb.create_sheet("A1 Demand Forecast")
    ws.append(["Hour", "Forecasted Defects"])
    for i in range(1, 25):
        ws.append([i, i * 0.5])

    # A2 Anomaly Detection
    ws = wb.create_sheet("A2 Anomaly Detection")
    ws.append(["Coil ID", "Anomaly Score"])
    ws.append(["CGL-001", 0.12])
    ws.append(["CGL-002", 0.95])
    ws.append(["CAL-001", 0.03])
    ws.append(["RCL-001", 0.45])

    # A3 RL Policy
    ws = wb.create_sheet("A3 RL Policy")
    ws.append(["Fatigue Level", "Time on line (min)", "Recommended Action"])
    ws.append(["Low (1-3)", "0-10", "Continue"])
    ws.append(["Medium (4-6)", "10-20", "Continue"])
    ws.append(["High (7-8)", "20-30", "Rotate"])
    ws.append(["Critical (9-10)", "30+", "Rotate Immediately"])

    # A4 Fatigue Predict
    ws = wb.create_sheet("A4 Fatigue Predict")
    ws.append(["Hour", "Predicted Fatigue (1-10)"])
    for i in range(1, 13):
        ws.append([i, round(5 + i * 0.2, 1)])

    # A5 DP Scheduling
    ws = wb.create_sheet("A5 DP Scheduling")
    ws.append(["Shift", "Inspectors Required"])
    ws.append(["Morning", 4])
    ws.append(["Afternoon", 3])
    ws.append(["Night", 2])

    # A6 Genetic Algorithm
    ws = wb.create_sheet("A6 Genetic Algorithm")
    ws.append(["Line", "Assigned Inspectors", "Optimal?"])
    ws.append(["CGL", 2, "Yes"])
    ws.append(["CAL", 2, "Yes"])
    ws.append(["RCL", 2, "Yes"])

    # A7 CUSUM Control
    ws = wb.create_sheet("A7 CUSUM Control")
    ws.append(["Sample", "CUSUM Statistic"])
    for i in range(1, 25):
        ws.append([i, round(i * 0.1, 2)])

    # A8 Monte Carlo
    ws = wb.create_sheet("A8 Monte Carlo")
    ws.append(["Risk Category", "Probability"])
    ws.append(["Low Risk", 0.70])
    ws.append(["Medium Risk", 0.20])
    ws.append(["High Risk", 0.10])

    # A9 Markov Chain
    ws = wb.create_sheet("A9 Markov Chain")
    ws.append(["Inspector State", "Steady-State Probability (%)"])
    ws.append(["Active", 55])
    ws.append(["Fatigued", 25])
    ws.append(["Rotating", 15])
    ws.append(["Absent", 3])
    ws.append(["Training", 2])

    # A10 Live Dashboard
    ws = wb.create_sheet("A10 Live Dashboard")
    ws.append(["Line", "OEE (%)", "Utilization (%)", "Alerts"])
    ws.append(["CGL", 87.5, 92.0, "OK"])
    ws.append(["CAL", 91.2, 88.5, "OK"])
    ws.append(["RCL", 79.3, 85.0, "Fatigue Alert"])

    return wb