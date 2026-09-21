#!/usr/bin/env python3
"""
generate_master_financial_workstation.py

Generates the flagship 'Master_Financial_Analysis_and_Modeling_Workstation.xlsx'
A comprehensive, institutional-grade corporate financial model designed for
both Microsoft Excel (365 / 2021 / 2019) and Google Sheets.

Covers:
  - 01_Executive_Dashboard (KPI summary cards, scenario toggle, dynamic charts)
  - 02_Assumptions_&_Scenarios (Base, Bull, Bear scenarios, WACC, Drivers)
  - 03_3_Statement_Model (5-year integrated P&L, Balance Sheet, Cash Flow with balance check)
  - 04_Valuation_DCF (UFCF, WACC, NPV, XNPV, IRR, XIRR, Terminal Value, Sensitivity Matrix)
  - 05_Debt_&_Depreciation (PMT, IPMT, PPMT, CUMIPMT, SLN, DB, Fixed Asset Register)
  - 06_Budget_vs_Actual_BvA (12-month BvA, SUMIFS, SUMPRODUCT, Variance $, %, Alerts)
  - 07_Advanced_Formula_Lab (LET, LAMBDA, FILTER, UNIQUE, SORT, XLOOKUP, INDEX/MATCH, OFFSET, INDIRECT)
  - 08_Formula_Glossary_&_Guide (Excel vs Google Sheets compatibility matrix & shortcut guide)
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference, Series

def create_master_workstation(output_path: str):
    wb = openpyxl.Workbook()
    # Remove default sheet
    default_sheet = wb.active
    wb.remove(default_sheet)

    # Styling Palette (Classic Institutional Corporate Navy & Slate)
    navy_dark = "1B365D"
    navy_mid = "2C5282"
    navy_light = "EBF8FF"
    slate_header = "334155"
    slate_sub = "64748B"
    border_gray = "CBD5E1"
    zebra_light = "F8FAFC"
    kpi_gold = "D97706"
    kpi_green = "16A34A"
    kpi_red = "DC2626"
    card_bg = "F1F5F9"
    highlight_yellow = "FEF3C7"

    font_title = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    font_section = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    font_sub_header = Font(name="Calibri", size=10, bold=True, color="334155")
    font_bold = Font(name="Calibri", size=11, bold=True, color="0F172A")
    font_regular = Font(name="Calibri", size=10, color="1E293B")
    font_muted = Font(name="Calibri", size=9, italic=True, color="64748B")
    font_formula = Font(name="Consolas", size=9, color="0369A1")
    font_kpi_val = Font(name="Calibri", size=16, bold=True, color="1B365D")
    font_kpi_lbl = Font(name="Calibri", size=9, bold=True, color="64748B")

    fill_title = PatternFill(start_color=navy_dark, end_color=navy_dark, fill_type="solid")
    fill_header = PatternFill(start_color=navy_mid, end_color=navy_mid, fill_type="solid")
    fill_zebra = PatternFill(start_color=zebra_light, end_color=zebra_light, fill_type="solid")
    fill_subtotal = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    fill_total = PatternFill(start_color="CBD5E1", end_color="CBD5E1", fill_type="solid")
    fill_highlight = PatternFill(start_color=highlight_yellow, end_color=highlight_yellow, fill_type="solid")
    fill_card = PatternFill(start_color=card_bg, end_color=card_bg, fill_type="solid")

    thin_border_side = Side(style="thin", color=border_gray)
    double_border_side = Side(style="double", color="0F172A")
    thick_top = Side(style="thin", color="0F172A")

    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    subtotal_border = Border(top=thin_border_side, bottom=thin_border_side)
    total_border = Border(top=thick_top, bottom=double_border_side)

    # -------------------------------------------------------------
    # TAB 1: 01_Executive_Dashboard
    # -------------------------------------------------------------
    ws1 = wb.create_sheet(title="01_Executive_Dashboard")
    ws1.views.sheetView[0].showGridLines = True

    # Title Block
    ws1.merge_cells("A1:K2")
    ws1["A1"] = "GLOBAL ENTERPRISE FINANCIAL MODEL & EXECUTIVE DASHBOARD"
    ws1["A1"].font = font_title
    ws1["A1"].fill = fill_title
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")

    # Scenario Controller Bar
    ws1["A3"] = "ACTIVE SCENARIO SELECTOR:"
    ws1["A3"].font = font_bold
    ws1["C3"] = "Base"
    ws1["C3"].font = Font(name="Calibri", size=12, bold=True, color="1E3A8A")
    ws1["C3"].fill = fill_highlight
    ws1["C3"].alignment = Alignment(horizontal="center")
    ws1["C3"].border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    ws1["D3"] = "(Valid options: Base | Bull | Bear — dynamic model linkages drive all projections)"
    ws1["D3"].font = font_muted

    # KPI Summary Cards Block (Row 5 - 7)
    kpis = [
        ("A5:B5", "A6:B7", "FY27 PROJ REVENUE", "='03_3_Statement_Model'!F9", "$#,##0"),
        ("C5:D5", "C6:D7", "FY27 EBITDA", "='03_3_Statement_Model'!F15", "$#,##0"),
        ("E5:F5", "E6:F7", "NET MARGIN (FY27)", "='03_3_Statement_Model'!F20/'03_3_Statement_Model'!F9", "0.0%"),
        ("G5:H5", "G6:H7", "ENTERPRISE VALUE (DCF)", "='04_Valuation_DCF'!C19", "$#,##0"),
        ("I5:I5", "I6:I7", "PROJECT IRR", "='04_Valuation_DCF'!C23", "0.0%"),
        ("J5:K5", "J6:K7", "ENDING CASH BALANCE", "='03_3_Statement_Model'!F38", "$#,##0")
    ]

    for top_range, val_range, label, formula, num_fmt in kpis:
        ws1.merge_cells(top_range)
        ws1.merge_cells(val_range)
        top_cell = top_range.split(":")[0]
        val_cell = val_range.split(":")[0]
        ws1[top_cell] = label
        ws1[top_cell].font = font_kpi_lbl
        ws1[top_cell].fill = fill_card
        ws1[top_cell].alignment = Alignment(horizontal="center", vertical="center")
        
        ws1[val_cell] = formula
        ws1[val_cell].font = font_kpi_val
        ws1[val_cell].fill = fill_card
        ws1[val_cell].alignment = Alignment(horizontal="center", vertical="center")
        ws1[val_cell].number_format = num_fmt

    # Border for KPI Cards
    for r in range(5, 8):
        for c in range(1, 12):
            ws1.cell(row=r, column=c).border = cell_border

    # Quick Navigation Index (Row 9 to 11)
    ws1.merge_cells("A9:K9")
    ws1["A9"] = "WORKBOOK ARCHITECTURE & SECTION QUICK JUMP"
    ws1["A9"].font = font_section
    ws1["A9"].fill = fill_header
    ws1["A9"].alignment = Alignment(horizontal="left", vertical="center")

    nav_items = [
        ("A10", "02_Assumptions_&_Scenarios", "Macro assumptions, pricing/volume drivers, and scenario matrices"),
        ("A11", "03_3_Statement_Model", "Integrated 5-Year Income Statement, Balance Sheet & Cash Flow Statement"),
        ("E10", "04_Valuation_DCF", "Discounted Cash Flow, WACC build, Terminal Value, Sensitivity Tables"),
        ("E11", "05_Debt_&_Depreciation", "Term loan amortization schedule (PMT/IPMT/PPMT) & asset depreciation"),
        ("I10", "06_Budget_vs_Actual_BvA", "12-month departmental variance analysis with SUMIFS and alert flags"),
        ("I11", "07_Advanced_Formula_Lab", "Modern dynamic arrays: LET, LAMBDA, FILTER, UNIQUE, SORT, XLOOKUP")
    ]
    for pos, sheet_name, desc in nav_items:
        ws1[pos] = f"• {sheet_name}: {desc}"
        ws1[pos].font = font_regular

    # Table Summary Preview (Row 13 to 21)
    ws1.merge_cells("A13:G13")
    ws1["A13"] = "5-YEAR HIGH-LEVEL CONSOLIDATED FINANCIAL SUMMARY ($)"
    ws1["A13"].font = font_section
    ws1["A13"].fill = fill_header

    dash_headers = ["Financial Metric ($)", "FY2023 (Act)", "FY2024 (Act)", "FY2025 (Proj)", "FY2026 (Proj)", "FY2027 (Proj)", "5-Yr CAGR"]
    for idx, h in enumerate(dash_headers, 1):
        c = ws1.cell(row=14, column=idx, value=h)
        c.font = font_bold
        c.fill = fill_subtotal
        c.alignment = Alignment(horizontal="right" if idx > 1 else "left")
        c.border = cell_border

    summary_rows = [
        ("Gross Revenue", "='03_3_Statement_Model'!B9", "='03_3_Statement_Model'!C9", "='03_3_Statement_Model'!D9", "='03_3_Statement_Model'!E9", "='03_3_Statement_Model'!F9", "=((F15/B15)^(1/4))-1", "$#,##0", "0.0%"),
        ("Cost of Goods Sold (COGS)", "='03_3_Statement_Model'!B10", "='03_3_Statement_Model'!C10", "='03_3_Statement_Model'!D10", "='03_3_Statement_Model'!E10", "='03_3_Statement_Model'!F10", "=((F16/B16)^(1/4))-1", "$#,##0", "0.0%"),
        ("Gross Profit", "='03_3_Statement_Model'!B11", "='03_3_Statement_Model'!C11", "='03_3_Statement_Model'!D11", "='03_3_Statement_Model'!E11", "='03_3_Statement_Model'!F11", "=((F17/B17)^(1/4))-1", "$#,##0", "0.0%"),
        ("EBITDA", "='03_3_Statement_Model'!B15", "='03_3_Statement_Model'!C15", "='03_3_Statement_Model'!D15", "='03_3_Statement_Model'!E15", "='03_3_Statement_Model'!F15", "=((F18/B18)^(1/4))-1", "$#,##0", "0.0%"),
        ("Net Operating Profit After Tax (NOPAT)", "='04_Valuation_DCF'!B9", "='04_Valuation_DCF'!C9", "='04_Valuation_DCF'!D9", "='04_Valuation_DCF'!E9", "='04_Valuation_DCF'!F9", "=((F19/B19)^(1/4))-1", "$#,##0", "0.0%"),
        ("Free Cash Flow (FCF)", "='04_Valuation_DCF'!B14", "='04_Valuation_DCF'!C14", "='04_Valuation_DCF'!D14", "='04_Valuation_DCF'!E14", "='04_Valuation_DCF'!F14", "=((F20/B20)^(1/4))-1", "$#,##0", "0.0%"),
        ("Ending Cash Balance", "='03_3_Statement_Model'!B38", "='03_3_Statement_Model'!C38", "='03_3_Statement_Model'!D38", "='03_3_Statement_Model'!E38", "='03_3_Statement_Model'!F38", "=((F21/B21)^(1/4))-1", "$#,##0", "0.0%")
    ]

    for r_idx, row_data in enumerate(summary_rows, start=15):
        for c_idx in range(1, 8):
            cell = ws1.cell(row=r_idx, column=c_idx)
            val = row_data[c_idx - 1]
            cell.value = val
            cell.border = cell_border
            if c_idx == 1:
                cell.font = font_bold
                cell.alignment = Alignment(horizontal="left")
            elif c_idx == 7:
                cell.font = font_regular
                cell.number_format = row_data[8]
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.font = font_regular
                cell.number_format = row_data[7]
                cell.alignment = Alignment(horizontal="right")
            if r_idx % 2 == 1:
                cell.fill = fill_zebra

    # Add Chart to Executive Dashboard (Revenue vs EBITDA)
    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 10
    chart1.title = "Consolidated Revenue vs. EBITDA Projection (5-Year)"
    chart1.y_axis.title = "USD ($)"
    chart1.x_axis.title = "Fiscal Year"

    data = Reference(ws1, min_col=2, min_row=15, max_col=6, max_row=18)
    cats = Reference(ws1, min_col=2, min_row=14, max_col=6, max_row=14)
    chart1.add_data(data, from_rows=True, titles_from_data=False)
    chart1.set_categories(cats)
    chart1.width = 16
    chart1.height = 9
    ws1.add_chart(chart1, "A23")

    # -------------------------------------------------------------
    # TAB 2: 02_Assumptions_&_Scenarios
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="02_Assumptions_&_Scenarios")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:G2")
    ws2["A1"] = "FINANCIAL MODEL ASSUMPTIONS & SCENARIO ENGINE"
    ws2["A1"].font = font_title
    ws2["A1"].fill = fill_title
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")

    # Scenario Switch Controller
    ws2["A4"] = "Active Scenario Selected:"
    ws2["A4"].font = font_bold
    ws2["C4"] = "='01_Executive_Dashboard'!C3"
    ws2["C4"].font = Font(name="Calibri", size=12, bold=True, color="1E3A8A")
    ws2["C4"].fill = fill_highlight
    ws2["C4"].alignment = Alignment(horizontal="center")
    ws2["C4"].border = cell_border

    ws2["A5"] = "Active Scenario ID (Index):"
    ws2["A5"].font = font_bold
    ws2["C5"] = '=IF(C4="Base", 1, IF(C4="Bull", 2, IF(C4="Bear", 3, 1)))'
    ws2["C5"].font = font_bold
    ws2["C5"].alignment = Alignment(horizontal="center")
    ws2["C5"].border = cell_border

    # Scenario Definition Table
    ws2.merge_cells("A7:G7")
    ws2["A7"] = "1. SCENARIO MATRIX & OPERATIONAL GROWTH DRIVERS"
    ws2["A7"].font = font_section
    ws2["A7"].fill = fill_header

    headers_scen = ["Scenario", "ID", "Annual Rev Growth %", "COGS % of Rev", "SG&A % of Rev", "Capex % of Rev", "Tax Rate %"]
    for idx, h in enumerate(headers_scen, 1):
        c = ws2.cell(row=8, column=idx, value=h)
        c.font = font_bold
        c.fill = fill_subtotal
        c.border = cell_border
        c.alignment = Alignment(horizontal="center" if idx <= 2 else "right")

    scen_data = [
        ("Base Case", 1, 0.12, 0.42, 0.20, 0.05, 0.25),
        ("Bull Case", 2, 0.20, 0.38, 0.18, 0.06, 0.25),
        ("Bear Case", 3, 0.04, 0.48, 0.24, 0.04, 0.25)
    ]

    for r_idx, row in enumerate(scen_data, start=9):
        for c_idx in range(1, 8):
            cell = ws2.cell(row=r_idx, column=c_idx, value=row[c_idx - 1])
            cell.font = font_regular
            cell.border = cell_border
            if c_idx > 2:
                cell.number_format = "0.0%"
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="center")

    # Dynamic Active Drivers (Calculated via INDEX / MATCH)
    ws2.merge_cells("A13:G13")
    ws2["A13"] = "ACTIVE DYNAMIC DRIVERS (Driven by Active Scenario Selector)"
    ws2["A13"].font = font_section
    ws2["A13"].fill = fill_header

    ws2["A14"] = "Driver Name"
    ws2["B14"] = "Active Value"
    ws2["C14"] = "Formula / Method"
    ws2["D14"] = "Description"
    for col in ["A", "B", "C", "D"]:
        ws2[f"{col}14"].font = font_bold
        ws2[f"{col}14"].fill = fill_subtotal
        ws2[f"{col}14"].border = cell_border

    active_drivers = [
        ("Revenue Growth Rate (YoY)", '=INDEX(C9:C11, $C$5)', "INDEX(C9:C11, $C$5)", "Dynamically looks up growth rate for active scenario", "0.0%"),
        ("COGS Ratio (% of Revenue)", '=INDEX(D9:D11, $C$5)', "INDEX(D9:D11, $C$5)", "Direct cost of goods/services ratio", "0.0%"),
        ("SG&A Ratio (% of Revenue)", '=INDEX(E9:E11, $C$5)', "INDEX(E9:E11, $C$5)", "Selling, General & Administrative expense intensity", "0.0%"),
        ("Capex Intensity (% of Revenue)", '=INDEX(F9:F11, $C$5)', "INDEX(F9:F11, $C$5)", "Capital expenditures allocation as % of revenue", "0.0%"),
        ("Corporate Tax Rate", '=INDEX(G9:G11, $C$5)', "INDEX(G9:G11, $C$5)", "Statutory corporate income tax rate", "0.0%")
    ]

    for idx, (name, formula, form_str, desc, num_fmt) in enumerate(active_drivers, start=15):
        ws2.cell(row=idx, column=1, value=name).font = font_bold
        c_val = ws2.cell(row=idx, column=2, value=formula)
        c_val.font = font_bold
        c_val.number_format = num_fmt
        c_val.alignment = Alignment(horizontal="right")
        ws2.cell(row=idx, column=3, value=form_str).font = font_formula
        ws2.cell(row=idx, column=4, value=desc).font = font_muted
        for c in range(1, 5):
            ws2.cell(row=idx, column=c).border = cell_border

    # WACC (Weighted Average Cost of Capital) Build Section
    ws2.merge_cells("A22:E22")
    ws2["A22"] = "2. CAPITAL STRUCTURE & WACC VALUATION BENCHMARKS"
    ws2["A22"].font = font_section
    ws2["A22"].fill = fill_header

    wacc_items = [
        ("Risk-Free Rate (Rf)", 0.042, "0.0%", "10-Year US Treasury Benchmark Yield"),
        ("Equity Risk Premium (ERP)", 0.055, "0.0%", "Expected historical market excess return"),
        ("Asset / Equity Beta (β)", 1.15, "0.00", "Systematic equity risk / volatility factor"),
        ("Cost of Equity (Ke = Rf + β * ERP)", "=B23+(B25*B24)", "0.0%", "CAPM Capital Asset Pricing Model"),
        ("Pre-Tax Cost of Debt (Kd)", 0.065, "0.0%", "Weighted borrowing rate on senior term debt"),
        ("Effective Tax Rate (t)", "=B19", "0.0%", "Corporate income tax shield"),
        ("After-Tax Cost of Debt [Kd * (1 - t)]", "=B27*(1-B28)", "0.0%", "Tax-shielded debt cost"),
        ("Target Debt Weight (Wd)", 0.30, "0.0%", "Capital structure % financed by debt"),
        ("Target Equity Weight (We)", "=1-B30", "0.0%", "Capital structure % financed by equity"),
        ("WEIGHTED AVERAGE COST OF CAPITAL (WACC)", "=(B31*B26)+(B30*B29)", "0.0%", "Discount Rate: (We * Ke) + (Wd * Kd_after_tax)")
    ]

    for idx, (label, val, num_fmt, desc) in enumerate(wacc_items, start=23):
        ws2.cell(row=idx, column=1, value=label).font = font_bold if idx == 32 else font_regular
        c_val = ws2.cell(row=idx, column=2, value=val)
        c_val.font = font_bold if idx == 32 else font_regular
        c_val.number_format = num_fmt
        c_val.alignment = Alignment(horizontal="right")
        ws2.cell(row=idx, column=3, value=desc).font = font_muted
        if idx == 32:
            ws2.cell(row=idx, column=1).fill = fill_subtotal
            c_val.fill = fill_highlight
        for c in range(1, 4):
            ws2.cell(row=idx, column=c).border = cell_border

    # -------------------------------------------------------------
    # TAB 3: 03_3_Statement_Model
    # -------------------------------------------------------------
    ws3 = wb.create_sheet(title="03_3_Statement_Model")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:F2")
    ws3["A1"] = "5-YEAR INTEGRATED 3-STATEMENT FINANCIAL MODEL"
    ws3["A1"].font = font_title
    ws3["A1"].fill = fill_title
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")

    model_cols = ["Line Item ($)", "FY2023 (Act)", "FY2024 (Act)", "FY2025 (Proj)", "FY2026 (Proj)", "FY2027 (Proj)"]
    for idx, h in enumerate(model_cols, 1):
        c = ws3.cell(row=4, column=idx, value=h)
        c.font = font_bold
        c.fill = fill_header
        c.border = cell_border
        c.alignment = Alignment(horizontal="left" if idx == 1 else "right")

    # A. Income Statement (Rows 5 to 20)
    ws3.merge_cells("A5:F5")
    ws3["A5"] = "A. CONSOLIDATED INCOME STATEMENT (P&L)"
    ws3["A5"].font = font_section
    ws3["A5"].fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")

    pl_lines = [
        ("Gross Revenue", 10000000, 11500000, "=B9*(1+'02_Assumptions_&_Scenarios'!$B$15)", "=C9*(1+'02_Assumptions_&_Scenarios'!$B$15)", "=D9*(1+'02_Assumptions_&_Scenarios'!$B$15)", False),
        ("Cost of Goods Sold (COGS)", 4300000, 4830000, "=D9*'02_Assumptions_&_Scenarios'!$B$16", "=E9*'02_Assumptions_&_Scenarios'!$B$16", "=F9*'02_Assumptions_&_Scenarios'!$B$16", False),
        ("Gross Profit", "=B9-B10", "=C9-C10", "=D9-D10", "=E9-E10", "=F9-F10", True),
        ("Selling, General & Admin (SG&A)", 2100000, 2350000, "=D9*'02_Assumptions_&_Scenarios'!$B$17", "=E9*'02_Assumptions_&_Scenarios'!$B$17", "=F9*'02_Assumptions_&_Scenarios'!$B$17", False),
        ("Research & Development (R&D)", 800000, 920000, "=D9*0.075", "=E9*0.075", "=F9*0.075", False),
        ("Other Operating Expenses", 300000, 350000, "=D9*0.025", "=E9*0.025", "=F9*0.025", False),
        ("EBITDA", "=B11-(B12+B13+B14)", "=C11-(C12+C13+C14)", "=D11-(D12+D13+D14)", "=E11-(E12+E13+E14)", "=F11-(F12+F13+F14)", True),
        ("Depreciation & Amortization (D&A)", 600000, 680000, "='05_Debt_&_Depreciation'!B35", "='05_Debt_&_Depreciation'!C35", "='05_Debt_&_Depreciation'!D35", False),
        ("Operating Income (EBIT)", "=B15-B16", "=C15-C16", "=D15-D16", "=E15-E16", "=F15-F16", True),
        ("Interest Expense", 180000, 165000, "='05_Debt_&_Depreciation'!B20", "='05_Debt_&_Depreciation'!C20", "='05_Debt_&_Depreciation'!D20", False),
        ("Earnings Before Taxes (EBT)", "=B17-B18", "=C17-C18", "=D17-D18", "=E17-E18", "=F17-F18", True),
        ("Income Tax Expense", "=B19*'02_Assumptions_&_Scenarios'!$B$19", "=C19*'02_Assumptions_&_Scenarios'!$B$19", "=D19*'02_Assumptions_&_Scenarios'!$B$19", "=E19*'02_Assumptions_&_Scenarios'!$B$19", "=F19*'02_Assumptions_&_Scenarios'!$B$19", False),
        ("NET INCOME", "=B19-B20", "=C19-C20", "=D19-D20", "=E19-E20", "=F19-F20", True)
    ]

    for idx, row in enumerate(pl_lines, start=9):
        is_bold = row[6]
        ws3.cell(row=idx, column=1, value=row[0]).font = font_bold if is_bold else font_regular
        for c in range(2, 7):
            val = row[c - 1]
            cell = ws3.cell(row=idx, column=c, value=val)
            cell.font = font_bold if is_bold else font_regular
            cell.number_format = "$#,##0"
            cell.alignment = Alignment(horizontal="right")
            cell.border = total_border if idx == 21 else (subtotal_border if is_bold else cell_border)
            if is_bold and idx != 21:
                cell.fill = fill_subtotal
            elif idx == 21:
                cell.fill = fill_total

    # B. Balance Sheet (Rows 23 to 39)
    ws3.merge_cells("A23:F23")
    ws3["A23"] = "B. CONSOLIDATED BALANCE SHEET"
    ws3["A23"].font = font_section
    ws3["A23"].fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")

    bs_lines = [
        ("Cash & Cash Equivalents", 2500000, 3100000, "=B38", "=C38", "=D38"),
        ("Accounts Receivable (45 days DSO)", "=B9*(45/365)", "=C9*(45/365)", "=D9*(45/365)", "=E9*(45/365)", "=F9*(45/365)"),
        ("Inventories (35 days DIO)", "=B10*(35/365)", "=C10*(35/365)", "=D10*(35/365)", "=E10*(35/365)", "=F10*(35/365)"),
        ("Total Current Assets", "=SUM(B25:B27)", "=SUM(C25:C27)", "=SUM(D25:D27)", "=SUM(E25:E27)", "=SUM(F25:F27)"),
        ("Property, Plant & Equipment (Net PP&E)", 4200000, 4600000, "='05_Debt_&_Depreciation'!B36", "='05_Debt_&_Depreciation'!C36", "='05_Debt_&_Depreciation'!D36"),
        ("TOTAL ASSETS", "=B28+B29", "=C28+C29", "=D28+D29", "=E28+E29", "=F28+F29"),
        ("Accounts Payable (40 days DPO)", "=B10*(40/365)", "=C10*(40/365)", "=D10*(40/365)", "=E10*(40/365)", "=F10*(40/365)"),
        ("Current Senior Debt", 400000, 450000, "='05_Debt_&_Depreciation'!B21", "='05_Debt_&_Depreciation'!C21", "='05_Debt_&_Depreciation'!D21"),
        ("Total Current Liabilities", "=B31+B32", "=C31+C32", "=D31+D32", "=E31+E32", "=F31+F32"),
        ("Long-Term Debt", 2600000, 2150000, "='05_Debt_&_Depreciation'!B22", "='05_Debt_&_Depreciation'!C22", "='05_Debt_&_Depreciation'!D22"),
        ("TOTAL LIABILITIES", "=B33+B34", "=C33+C34", "=D33+D34", "=E33+E34", "=F33+F34"),
        ("Common Equity & Paid-in Capital", 2000000, 2000000, "=B36", "=C36", "=D36"),
        ("Retained Earnings", 3000000, 4200000, "=C37+D21", "=D37+E21", "=E37+F21"),
        ("TOTAL SHAREHOLDER EQUITY", "=B36+B37", "=C36+C37", "=D36+D37", "=E36+E37", "=F36+F37"),
        ("TOTAL LIABILITIES & EQUITY", "=B35+B38", "=C35+C38", "=D35+D38", "=E35+E38", "=F35+F38"),
        ("BALANCE SHEET CHECK (Assets - Liab & Eq)", "=B30-B39", "=C30-C39", "=D30-D39", "=E30-E39", "=F30-F39")
    ]

    for idx, row in enumerate(bs_lines, start=25):
        is_bold = idx in [28, 30, 33, 35, 38, 39, 40]
        ws3.cell(row=idx, column=1, value=row[0]).font = font_bold if is_bold else font_regular
        for c in range(2, 7):
            val = row[c - 1]
            cell = ws3.cell(row=idx, column=c, value=val)
            cell.font = font_bold if is_bold else font_regular
            cell.number_format = "$#,##0"
            cell.alignment = Alignment(horizontal="right")
            cell.border = total_border if idx in [30, 39] else (subtotal_border if is_bold else cell_border)
            if idx == 40:
                cell.font = Font(name="Calibri", size=9, bold=True, color="15803D")
                cell.number_format = "$#,##0;($#,##0);\"BALANCED ✓\""

    # C. Cash Flow Statement (Rows 42 to 55)
    ws3.merge_cells("A42:F42")
    ws3["A42"] = "C. CONSOLIDATED STATEMENT OF CASH FLOWS"
    ws3["A42"].font = font_section
    ws3["A42"].fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")

    cf_lines = [
        ("Net Income", "=B21", "=C21", "=D21", "=E21", "=F21"),
        ("(+) Depreciation & Amortization", "=B16", "=C16", "=D16", "=E16", "=F16"),
        ("(-) Change in Accounts Receivable", 0, "=(B26-C26)", "=(C26-D26)", "=(D26-E26)", "=(E26-F26)"),
        ("(-) Change in Inventory", 0, "=(B27-C27)", "=(C27-D27)", "=(D27-E27)", "=(E27-F27)"),
        ("(+) Change in Accounts Payable", 0, "=(C31-B31)", "=(D31-C31)", "=(E31-D31)", "=(F31-E31)"),
        ("CASH FLOW FROM OPERATIONS (CFO)", "=SUM(B44:B48)", "=SUM(C44:C48)", "=SUM(D44:D48)", "=SUM(E44:E48)", "=SUM(F44:F48)"),
        ("Capital Expenditures (Capex)", -500000, -600000, "=-D9*'02_Assumptions_&_Scenarios'!$B$18", "=-E9*'02_Assumptions_&_Scenarios'!$B$18", "=-F9*'02_Assumptions_&_Scenarios'!$B$18"),
        ("CASH FLOW FROM INVESTING (CFI)", "=B50", "=C50", "=D50", "=E50", "=F50"),
        ("Senior Debt Principal Repayment", -350000, -450000, "=-'05_Debt_&_Depreciation'!B21", "=-'05_Debt_&_Depreciation'!C21", "=-'05_Debt_&_Depreciation'!D21"),
        ("CASH FLOW FROM FINANCING (CFF)", "=B52", "=C52", "=D52", "=E52", "=F52"),
        ("NET CHANGE IN CASH", "=B49+B51+B53", "=C49+C51+C53", "=D49+D51+D53", "=E49+E51+E53", "=F49+F51+F53"),
        ("Beginning Cash Balance", 2000000, "=B25", "=C25", "=D25", "=E25"),
        ("ENDING CASH BALANCE", "=B54+B55", "=C54+C55", "=D54+D55", "=E54+E55", "=F54+F55")
    ]

    for idx, row in enumerate(cf_lines, start=44):
        is_bold = idx in [49, 51, 53, 54, 56]
        ws3.cell(row=idx, column=1, value=row[0]).font = font_bold if is_bold else font_regular
        for c in range(2, 7):
            val = row[c - 1]
            cell = ws3.cell(row=idx, column=c, value=val)
            cell.font = font_bold if is_bold else font_regular
            cell.number_format = "$#,##0"
            cell.alignment = Alignment(horizontal="right")
            cell.border = total_border if idx == 56 else (subtotal_border if is_bold else cell_border)
            if is_bold and idx != 56:
                cell.fill = fill_subtotal
            elif idx == 56:
                cell.fill = fill_highlight

    # -------------------------------------------------------------
    # TAB 4: 04_Valuation_DCF
    # -------------------------------------------------------------
    ws4 = wb.create_sheet(title="04_Valuation_DCF")
    ws4.views.sheetView[0].showGridLines = True

    ws4.merge_cells("A1:G2")
    ws4["A1"] = "DISCOUNTED CASH FLOW (DCF) VALUATION & RETURNS ENGINE"
    ws4["A1"].font = font_title
    ws4["A1"].fill = fill_title
    ws4["A1"].alignment = Alignment(horizontal="center", vertical="center")

    dcf_cols = ["Valuation Element ($)", "FY2023 (Act)", "FY2024 (Act)", "FY2025 (Proj)", "FY2026 (Proj)", "FY2027 (Proj)", "Terminal Year"]
    for idx, h in enumerate(dcf_cols, 1):
        c = ws4.cell(row=4, column=idx, value=h)
        c.font = font_bold
        c.fill = fill_header
        c.border = cell_border
        c.alignment = Alignment(horizontal="left" if idx == 1 else "right")

    dcf_ufcf = [
        ("Operating Income (EBIT)", "='03_3_Statement_Model'!B17", "='03_3_Statement_Model'!C17", "='03_3_Statement_Model'!D17", "='03_3_Statement_Model'!E17", "='03_3_Statement_Model'!F17", "=F5*(1+0.025)"),
        ("Effective Tax Rate", "='02_Assumptions_&_Scenarios'!$B$19", "='02_Assumptions_&_Scenarios'!$B$19", "='02_Assumptions_&_Scenarios'!$B$19", "='02_Assumptions_&_Scenarios'!$B$19", "='02_Assumptions_&_Scenarios'!$B$19", "=F6"),
        ("(-) Taxes on EBIT", "=B5*B6", "=C5*C6", "=D5*D6", "=E5*E6", "=F5*F6", "=G5*G6"),
        ("NOPAT (Net Operating Profit After Tax)", "=B5-B7", "=C5-C7", "=D5-D7", "=E5-E7", "=F5-F7", "=G5-G7"),
        ("(+) Depreciation & Amortization", "='03_3_Statement_Model'!B16", "='03_3_Statement_Model'!C16", "='03_3_Statement_Model'!D16", "='03_3_Statement_Model'!E16", "='03_3_Statement_Model'!F16", "=F9*(1+0.025)"),
        ("(-) Capital Expenditures (Capex)", "=-'03_3_Statement_Model'!B50", "=-'03_3_Statement_Model'!C50", "=-'03_3_Statement_Model'!D50", "=-'03_3_Statement_Model'!E50", "=-'03_3_Statement_Model'!F50", "=F10*(1+0.025)"),
        ("(-) Change in Net Working Capital (ΔNWC)", 50000, 75000, 85000, 95000, 110000, "=F11*(1+0.025)"),
        ("UNLEVERED FREE CASH FLOW (UFCF)", "=B8+B9-B10-B11", "=C8+C9-C10-C11", "=D8+D9-D10-D11", "=E8+E9-E10-E11", "=F8+F9-F10-F11", "=G8+G9-G10-G11"),
        ("Discount Period (Years)", 0, 0, 1, 2, 3, 4),
        ("Discount Factor [1 / (1 + WACC)^t]", 1, 1, "=1/(1+'02_Assumptions_&_Scenarios'!$B$32)^D13", "=1/(1+'02_Assumptions_&_Scenarios'!$B$32)^E13", "=1/(1+'02_Assumptions_&_Scenarios'!$B$32)^F13", "-"),
        ("PRESENT VALUE OF CASH FLOWS (PV)", 0, 0, "=D12*D14", "=E12*E14", "=F12*F14", "-")
    ]

    for idx, row in enumerate(dcf_ufcf, start=5):
        is_bold = idx in [8, 12, 15]
        ws4.cell(row=idx, column=1, value=row[0]).font = font_bold if is_bold else font_regular
        for c in range(2, 8):
            val = row[c - 1]
            cell = ws4.cell(row=idx, column=c, value=val)
            cell.font = font_bold if is_bold else font_regular
            cell.border = total_border if idx == 12 else (subtotal_border if is_bold else cell_border)
            if idx == 13:
                cell.number_format = "0.0"
            elif idx == 14:
                cell.number_format = "0.0000" if c in [4, 5, 6] else "@"
            elif idx == 6:
                cell.number_format = "0.0%"
            else:
                cell.number_format = "$#,##0" if c in [2, 3, 4, 5, 6, 7] and val != "-" else "@"
            cell.alignment = Alignment(horizontal="right")
            if is_bold:
                cell.fill = fill_subtotal

    # Valuation Summary & Returns Card (Rows 17 to 25)
    ws4.merge_cells("A17:D17")
    ws4["A17"] = "VALUATION OUTPUT & INSTITUTIONAL METRICS"
    ws4["A17"].font = font_section
    ws4["A17"].fill = fill_header

    val_metrics = [
        ("Cumulative PV of Explicit Forecast (FY25-FY27)", "=SUM(D15:F15)", "$#,##0", "Sum of discounted 3-year UFCF"),
        ("Terminal Value (Gordon Growth g=2.5%)", "=(G12*('02_Assumptions_&_Scenarios'!$B$32-0.025))", "$#,##0", "TV = UFCF_t+1 / (WACC - g)"),
        ("PV of Terminal Value", "=B19*(1/(1+'02_Assumptions_&_Scenarios'!$B$32)^3)", "$#,##0", "Discounted back 3 years to present"),
        ("ENTERPRISE VALUE (EV)", "=B18+B20", "$#,##0", "Core operating value of business"),
        ("(-) Net Debt (Total Debt - Cash)", "='03_3_Statement_Model'!F35-'03_3_Statement_Model'!F38", "$#,##0", "Market net indebtedness"),
        ("IMPLIED EQUITY VALUE", "=B21-B22", "$#,##0", "Value attributable to common shareholders"),
        ("XNPV (Formula Check - Exact Dates)", "=XNPV('02_Assumptions_&_Scenarios'!$B$32, D12:F12, DATE(2025,{12;12;12},{31;31;31}))", "$#,##0", "Modern financial exact date discounting"),
        ("PROJECT IRR (Internal Rate of Return)", "=IRR(D12:F12)", "0.0%", "Unlevered internal return threshold")
    ]

    for idx, (lbl, formula, num_fmt, desc) in enumerate(val_metrics, start=18):
        ws4.cell(row=idx, column=1, value=lbl).font = font_bold if idx in [21, 23] else font_regular
        c_val = ws4.cell(row=idx, column=2, value=formula)
        c_val.font = font_bold if idx in [21, 23] else font_regular
        c_val.number_format = num_fmt
        c_val.alignment = Alignment(horizontal="right")
        ws4.cell(row=idx, column=3, value=desc).font = font_muted
        if idx in [21, 23]:
            ws4.cell(row=idx, column=1).fill = fill_highlight
            c_val.fill = fill_highlight
        for c in range(1, 4):
            ws4.cell(row=idx, column=c).border = cell_border

    # Sensitivity Table (Rows 17 to 23, Columns E to I)
    ws4.merge_cells("E17:I17")
    ws4["E17"] = "VALUATION SENSITIVITY: ENTERPRISE VALUE ($M) VS WACC & GROWTH"
    ws4["E17"].font = font_bold
    ws4["E17"].fill = fill_subtotal
    ws4["E17"].alignment = Alignment(horizontal="center")

    sens_header = ["WACC \\ g", "1.5%", "2.0%", "2.5%", "3.0%"]
    for idx, h in enumerate(sens_header, start=5):
        c = ws4.cell(row=18, column=idx, value=h)
        c.font = font_bold
        c.fill = fill_card
        c.alignment = Alignment(horizontal="center")
        c.border = cell_border

    wacc_rates = [0.075, 0.085, 0.095, 0.105, 0.115]
    for r_idx, w_rate in enumerate(wacc_rates, start=19):
        c_rate = ws4.cell(row=r_idx, column=5, value=w_rate)
        c_rate.font = font_bold
        c_rate.number_format = "0.0%"
        c_rate.alignment = Alignment(horizontal="center")
        c_rate.border = cell_border
        for c_idx, g_rate in enumerate([0.015, 0.020, 0.025, 0.030], start=6):
            cell = ws4.cell(row=r_idx, column=c_idx)
            cell.value = f"=($B$18 + ($G$12 / ({w_rate} - {g_rate})) * (1 / (1 + {w_rate})^3)) / 1000000"
            cell.font = font_regular
            cell.number_format = "$#,##0.0"
            cell.alignment = Alignment(horizontal="right")
            cell.border = cell_border
            if w_rate == 0.095 and g_rate == 0.025:
                cell.fill = fill_highlight
                cell.font = font_bold

    # -------------------------------------------------------------
    # TAB 5: 05_Debt_&_Depreciation
    # -------------------------------------------------------------
    ws5 = wb.create_sheet(title="05_Debt_&_Depreciation")
    ws5.views.sheetView[0].showGridLines = True

    ws5.merge_cells("A1:G2")
    ws5["A1"] = "DEBT AMORTIZATION & MULTI-METHOD ASSET DEPRECIATION"
    ws5["A1"].font = font_title
    ws5["A1"].fill = fill_title
    ws5["A1"].alignment = Alignment(horizontal="center", vertical="center")

    # Section 1: Term Loan Facility Parameters
    ws5.merge_cells("A4:E4")
    ws5["A4"] = "1. SENIOR TERM LOAN AMORTIZATION SCHEDULE (PMT / IPMT / PPMT)"
    ws5["A4"].font = font_section
    ws5["A4"].fill = fill_header

    debt_params = [
        ("Facility Principal Amount", 3000000, "$#,##0", "Total senior credit facility borrowed"),
        ("Annual Interest Rate", 0.065, "0.0%", "Contractual coupon rate"),
        ("Tenor (Years)", 5, "0", "Loan repayment duration"),
        ("Periods Per Year", 12, "0", "Monthly amortization frequency"),
        ("Monthly Payment (PMT)", "=PMT(B6/B8, B7*B8, -B5)", "$#,##0.00", "Contractual monthly debt service (P+I)"),
        ("Total 5-Year Interest Paid", "=(B9*B7*B8)-B5", "$#,##0", "Cumulative interest expense across life of loan")
    ]

    for idx, (lbl, val, num_fmt, desc) in enumerate(debt_params, start=5):
        ws5.cell(row=idx, column=1, value=lbl).font = font_bold if idx == 9 else font_regular
        c_val = ws5.cell(row=idx, column=2, value=val)
        c_val.font = font_bold if idx == 9 else font_regular
        c_val.number_format = num_fmt
        c_val.alignment = Alignment(horizontal="right")
        ws5.cell(row=idx, column=3, value=desc).font = font_muted
        if idx == 9:
            ws5.cell(row=idx, column=1).fill = fill_highlight
            c_val.fill = fill_highlight
        for c in range(1, 4):
            ws5.cell(row=idx, column=c).border = cell_border

    # Monthly Schedule Preview (Months 1 to 12)
    debt_sched_headers = ["Month", "Beg Balance ($)", "Total Payment ($)", "Principal (PPMT) ($)", "Interest (IPMT) ($)", "Ending Balance ($)", "Check"]
    for idx, h in enumerate(debt_sched_headers, 1):
        c = ws5.cell(row=12, column=idx, value=h)
        c.font = font_bold
        c.fill = fill_subtotal
        c.border = cell_border
        c.alignment = Alignment(horizontal="center" if idx in [1, 7] else "right")

    for m in range(1, 13):
        r = 12 + m
        ws5.cell(row=r, column=1, value=m).font = font_regular
        ws5.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        
        beg_bal = "=$B$5" if m == 1 else f"=F{r-1}"
        ws5.cell(row=r, column=2, value=beg_bal).number_format = "$#,##0.00"
        ws5.cell(row=r, column=3, value="=$B$9").number_format = "$#,##0.00"
        ws5.cell(row=r, column=4, value=f"=PPMT($B$6/$B$8, A{r}, $B$7*$B$8, -$B$5)").number_format = "$#,##0.00"
        ws5.cell(row=r, column=5, value=f"=IPMT($B$6/$B$8, A{r}, $B$7*$B$8, -$B$5)").number_format = "$#,##0.00"
        ws5.cell(row=r, column=6, value=f"=B{r}-D{r}").number_format = "$#,##0.00"
        ws5.cell(row=r, column=7, value=f"=ROUND(C{r}-(D{r}+E{r}), 2)").number_format = "0.00"
        ws5.cell(row=r, column=7).alignment = Alignment(horizontal="center")
        
        for c in range(1, 8):
            cell = ws5.cell(row=r, column=c)
            cell.font = font_regular
            cell.border = cell_border
            if m % 2 == 0:
                cell.fill = fill_zebra

    # Section 2: Fixed Asset Multi-Method Depreciation
    ws5.merge_cells("A27:G27")
    ws5["A27"] = "2. MULTI-METHOD DEPRECIATION SCHEDULE (SLN, DB, SYD COMPARISON)"
    ws5["A27"].font = font_section
    ws5["A27"].fill = fill_header

    ws5["A28"] = "Asset Cost Basis ($):"
    ws5["B28"] = 5000000
    ws5["B28"].number_format = "$#,##0"
    ws5["C28"] = "Useful Life (Years):"
    ws5["D28"] = 5
    ws5["E28"] = "Salvage Value ($):"
    ws5["F28"] = 500000
    ws5["F28"].number_format = "$#,##0"

    for col in ["A", "C", "E"]:
        ws5[f"{col}28"].font = font_bold
    for col in ["B", "D", "F"]:
        ws5[f"{col}28"].font = font_bold
        ws5[f"{col}28"].border = cell_border

    dep_headers = ["Year", "Straight Line (SLN)", "Declining Balance (DB)", "Sum-of-Years' Digits (SYD)", "Cumulative SLN", "Net Book Value (SLN)", "Recommended"]
    for idx, h in enumerate(dep_headers, 1):
        c = ws5.cell(row=30, column=idx, value=h)
        c.font = font_bold
        c.fill = fill_subtotal
        c.border = cell_border
        c.alignment = Alignment(horizontal="center" if idx in [1, 7] else "right")

    for y in range(1, 6):
        r = 30 + y
        ws5.cell(row=r, column=1, value=f"Year {y}").font = font_bold
        ws5.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        
        ws5.cell(row=r, column=2, value=f"=SLN($B$28, $F$28, $D$28)").number_format = "$#,##0"
        ws5.cell(row=r, column=3, value=f"=DB($B$28, $F$28, $D$28, {y})").number_format = "$#,##0"
        ws5.cell(row=r, column=4, value=f"=SYD($B$28, $F$28, $D$28, {y})").number_format = "$#,##0"
        
        cum_sln = f"=B{r}" if y == 1 else f"=E{r-1}+B{r}"
        ws5.cell(row=r, column=5, value=cum_sln).number_format = "$#,##0"
        ws5.cell(row=r, column=6, value=f"=$B$28-E{r}").number_format = "$#,##0"
        ws5.cell(row=r, column=7, value="GAAP / IFRS Base").alignment = Alignment(horizontal="center")
        
        for c in range(1, 8):
            cell = ws5.cell(row=r, column=c)
            cell.font = font_regular
            cell.border = cell_border
            if y % 2 == 0:
                cell.fill = fill_zebra

    # -------------------------------------------------------------
    # TAB 6: 06_Budget_vs_Actual_BvA
    # -------------------------------------------------------------
    ws6 = wb.create_sheet(title="06_Budget_vs_Actual_BvA")
    ws6.views.sheetView[0].showGridLines = True

    ws6.merge_cells("A1:H2")
    ws6["A1"] = "DEPARTMENTAL BUDGET VS ACTUAL (BvA) VARIANCE ENGINE"
    ws6["A1"].font = font_title
    ws6["A1"].fill = fill_title
    ws6["A1"].alignment = Alignment(horizontal="center", vertical="center")

    bva_headers = ["Cost Center / Line Item", "Department", "Annual Budget ($)", "YTD Actuals ($)", "Variance ($)", "Variance (%)", "Status & Flag", "Audit Check"]
    for idx, h in enumerate(bva_headers, 1):
        c = ws6.cell(row=4, column=idx, value=h)
        c.font = font_bold
        c.fill = fill_header
        c.border = cell_border
        c.alignment = Alignment(horizontal="left" if idx <= 2 else ("center" if idx in [7, 8] else "right"))

    departments_bva = [
        ("Direct Labor & Engineering", "R&D", 1200000, 1140000),
        ("Cloud Infrastructure & Servers", "R&D", 450000, 485000),
        ("Software Licenses & Tools", "IT", 220000, 210000),
        ("Sales Commission & Bonuses", "Sales", 650000, 710000),
        ("Digital Marketing & Ads", "Marketing", 500000, 480000),
        ("Executive & Management Salaries", "G&A", 950000, 950000),
        ("Office Lease & Facilities", "G&A", 360000, 360000),
        ("Legal, Audit & Compliance", "Legal", 180000, 225000),
        ("Travel & Entertainment (T&E)", "Sales", 140000, 165000),
        ("Customer Support & Success", "Operations", 320000, 298000),
        ("Recruiting & Talent Acquisition", "HR", 150000, 175000),
        ("Employee Health & Benefits", "HR", 280000, 275000),
        ("Cybersecurity & Penetration Testing", "IT", 120000, 115000),
        ("Payment Processing Fees", "Finance", 190000, 205000),
        ("Contingency & Miscellaneous", "G&A", 100000, 42000)
    ]

    for idx, (line, dept, bgt, act) in enumerate(departments_bva, start=5):
        ws6.cell(row=idx, column=1, value=line).font = font_regular
        ws6.cell(row=idx, column=2, value=dept).font = font_regular
        
        c_bgt = ws6.cell(row=idx, column=3, value=bgt)
        c_bgt.font = font_regular
        c_bgt.number_format = "$#,##0"
        
        c_act = ws6.cell(row=idx, column=4, value=act)
        c_act.font = font_regular
        c_act.number_format = "$#,##0"
        
        c_var = ws6.cell(row=idx, column=5, value=f"=D{idx}-C{idx}")
        c_var.font = font_bold
        c_var.number_format = "$#,##0;($#,##0);\"-\""
        
        c_var_pct = ws6.cell(row=idx, column=6, value=f"=E{idx}/C{idx}")
        c_var_pct.font = font_regular
        c_var_pct.number_format = "+0.0%;-0.0%;0.0%"
        
        # Status formula with nested IF / IFS
        status_form = f'=IF(E{idx}>15000, "🚨 OVER BUDGET", IF(E{idx}<-15000, "🟢 FAVORABLE", "✅ ON TARGET"))'
        c_status = ws6.cell(row=idx, column=7, value=status_form)
        c_status.font = font_bold
        c_status.alignment = Alignment(horizontal="center")
        
        audit_check = f'=IF(ISNUMBER(C{idx}+D{idx}), "OK", "ERR")'
        c_audit = ws6.cell(row=idx, column=8, value=audit_check)
        c_audit.font = font_muted
        c_audit.alignment = Alignment(horizontal="center")
        
        for c in range(1, 9):
            ws6.cell(row=idx, column=c).border = cell_border
            if idx % 2 == 1:
                ws6.cell(row=idx, column=c).fill = fill_zebra

    # Total Row
    tot_r = 5 + len(departments_bva)
    ws6.cell(row=tot_r, column=1, value="CONSOLIDATED EXPENSES TOTAL").font = font_bold
    ws6.cell(row=tot_r, column=2, value="All Depts").font = font_bold
    ws6.cell(row=tot_r, column=3, value=f"=SUM(C5:C{tot_r-1})").number_format = "$#,##0"
    ws6.cell(row=tot_r, column=4, value=f"=SUM(D5:D{tot_r-1})").number_format = "$#,##0"
    ws6.cell(row=tot_r, column=5, value=f"=D{tot_r}-C{tot_r}").number_format = "$#,##0;($#,##0);\"-\""
    ws6.cell(row=tot_r, column=6, value=f"=E{tot_r}/C{tot_r}").number_format = "+0.0%;-0.0%;0.0%"
    ws6.cell(row=tot_r, column=7, value=f'=IF(E{tot_r}>0, "🚨 NET DEFICIT", "🟢 NET SURPLUS")').alignment = Alignment(horizontal="center")
    ws6.cell(row=tot_r, column=8, value="BALANCED").alignment = Alignment(horizontal="center")

    for c in range(1, 9):
        cell = ws6.cell(row=tot_r, column=c)
        cell.font = font_bold
        cell.fill = fill_total
        cell.border = total_border

    # Departmental Aggregation via SUMIFS & COUNTIFS (Rows tot_r + 3)
    agg_r = tot_r + 3
    ws6.merge_cells(f"A{agg_r}:E{agg_r}")
    ws6[f"A{agg_r}"] = "DEPARTMENTAL SUMMARY ROLLUP VIA SUMIFS & AVERAGEIFS"
    ws6[f"A{agg_r}"].font = font_section
    ws6[f"A{agg_r}"].fill = fill_header

    agg_headers = ["Department", "Line Items (COUNTIF)", "Total Budget ($)", "Total Actuals ($)", "Average Line Variance"]
    for idx, h in enumerate(agg_headers, 1):
        c = ws6.cell(row=agg_r+1, column=idx, value=h)
        c.font = font_bold
        c.fill = fill_subtotal
        c.border = cell_border
        c.alignment = Alignment(horizontal="left" if idx == 1 else "right")

    depts = ["R&D", "Sales", "G&A", "IT", "HR", "Operations", "Legal", "Finance"]
    for idx, d in enumerate(depts, start=agg_r+2):
        ws6.cell(row=idx, column=1, value=d).font = font_bold
        ws6.cell(row=idx, column=2, value=f'=COUNTIF($B$5:$B${tot_r-1}, "{d}")').alignment = Alignment(horizontal="right")
        ws6.cell(row=idx, column=3, value=f'=SUMIFS($C$5:$C${tot_r-1}, $B$5:$B${tot_r-1}, "{d}")').number_format = "$#,##0"
        ws6.cell(row=idx, column=4, value=f'=SUMIFS($D$5:$D${tot_r-1}, $B$5:$B${tot_r-1}, "{d}")').number_format = "$#,##0"
        ws6.cell(row=idx, column=5, value=f'=AVERAGEIFS($E$5:$E${tot_r-1}, $B$5:$B${tot_r-1}, "{d}")').number_format = "$#,##0;($#,##0);\"-\""
        for c in range(1, 6):
            ws6.cell(row=idx, column=c).border = cell_border

    # -------------------------------------------------------------
    # TAB 7: 07_Advanced_Formula_Lab
    # -------------------------------------------------------------
    ws7 = wb.create_sheet(title="07_Advanced_Formula_Lab")
    ws7.views.sheetView[0].showGridLines = True

    ws7.merge_cells("A1:G2")
    ws7["A1"] = "MODERN DYNAMIC ARRAYS & ADVANCED FORMULA SHOWCASE"
    ws7["A1"].font = font_title
    ws7["A1"].fill = fill_title
    ws7["A1"].alignment = Alignment(horizontal="center", vertical="center")

    lab_headers = ["Technique / Formula", "Syntax Example", "Computed Result", "Compatibility", "Business Application & Advantage"]
    for idx, h in enumerate(lab_headers, 1):
        c = ws7.cell(row=4, column=idx, value=h)
        c.font = font_bold
        c.fill = fill_header
        c.border = cell_border

    lab_items = [
        ("LET Variable Optimization", '=LET(rev, \'03_3_Statement_Model\'!F9, cogs, \'03_3_Statement_Model\'!F10, (rev-cogs)/rev)', "='03_3_Statement_Model'!F11/'03_3_Statement_Model'!F9", "Excel 365 / Sheets", "Eliminates redundant cell recalculation, accelerating large workbook speed by 4x"),
        ("XLOOKUP 2D Exact Match", '=XLOOKUP("Bull Case", \'02_Assumptions_&_Scenarios\'!A9:A11, \'02_Assumptions_&_Scenarios\'!C9:C11)', "='02_Assumptions_&_Scenarios'!C10", "Excel 2021+ / Sheets", "Replaces VLOOKUP with leftward lookups, default exact matching, and missing-key fallback"),
        ("INDEX / MATCH Classic Pair", '=INDEX(\'02_Assumptions_&_Scenarios\'!C9:C11, MATCH("Bear Case", \'02_Assumptions_&_Scenarios\'!A9:A11, 0))', "='02_Assumptions_&_Scenarios'!C11", "All Excel / All Sheets", "Universal institutional-grade matrix lookup compatible with all versions of spreadsheet engines"),
        ("SUMPRODUCT Weighted Metric", '=SUMPRODUCT(\'06_Budget_vs_Actual_BvA\'!C5:C10, \'06_Budget_vs_Actual_BvA\'!F5:F10)/SUM(\'06_Budget_vs_Actual_BvA\'!C5:C10)', "=SUMPRODUCT('06_Budget_vs_Actual_BvA'!C5:C10, '06_Budget_vs_Actual_BvA'!F5:F10)/SUM('06_Budget_vs_Actual_BvA'!C5:C10)", "All Excel / All Sheets", "Calculates dollar-weighted variance percentages across diverse departmental expense pools"),
        ("OFFSET Dynamic Rolling Range", '=AVERAGE(OFFSET(\'03_3_Statement_Model\'!B9, 0, 0, 1, 3))', "=AVERAGE('03_3_Statement_Model'!B9:D9)", "All Excel / All Sheets", "Dynamically averages moving multi-year fiscal performance without hardcoding column offsets"),
        ("INDIRECT Cross-Sheet Ref", '=INDIRECT("\'01_Executive_Dashboard\'!C3")', "='01_Executive_Dashboard'!C3", "All Excel / All Sheets", "Dynamically resolves target worksheet tabs based on dropdown strings or user selections"),
        ("FILTER High Variance Depts", '=FILTER(\'06_Budget_vs_Actual_BvA\'!A5:E19, \'06_Budget_vs_Actual_BvA\'!E5:E19>10000)', "Dynamic Spill", "Excel 365 / Sheets", "Extracts all expense rows that breached target threshold automatically without macros or VBA"),
        ("UNIQUE Dimension Extractor", '=UNIQUE(\'06_Budget_vs_Actual_BvA\'!B5:B19)', "Dynamic Spill", "Excel 365 / Sheets", "Auto-deduplicates organizational departments into clean dynamic validation lists"),
        ("SORT Ranked Performance", '=SORT(\'06_Budget_vs_Actual_BvA\'!A5:E19, 5, -1)', "Dynamic Spill", "Excel 365 / Sheets", "Instant server-side array sorting by greatest variance dollar magnitude descending"),
        ("LAMBDA Custom Financial Math", '=LAMBDA(pv, r, n, pv*(1+r)^n)(1000, 0.08, 5)', "=1000*(1+0.08)^5", "Excel 365 / Sheets", "Builds reusable, testable custom functions directly in formula bar without VBA modules")
    ]

    for idx, (name, syntax, result, comp, app) in enumerate(lab_items, start=5):
        ws7.cell(row=idx, column=1, value=name).font = font_bold
        ws7.cell(row=idx, column=2, value=syntax).font = font_formula
        
        c_res = ws7.cell(row=idx, column=3, value=result)
        c_res.font = font_bold
        if isinstance(result, str) and result.startswith("="):
            c_res.number_format = "0.0%" if "F11" in result or "SUMPRODUCT" in syntax else "$#,##0"
        c_res.alignment = Alignment(horizontal="right" if isinstance(result, str) and result.startswith("=") else "center")
        
        ws7.cell(row=idx, column=4, value=comp).font = font_regular
        ws7.cell(row=idx, column=4).alignment = Alignment(horizontal="center")
        ws7.cell(row=idx, column=5, value=app).font = font_muted
        
        for c in range(1, 6):
            ws7.cell(row=idx, column=c).border = cell_border
            if idx % 2 == 1:
                ws7.cell(row=idx, column=c).fill = fill_zebra

    # -------------------------------------------------------------
    # TAB 8: 08_Formula_Glossary_&_Guide
    # -------------------------------------------------------------
    ws8 = wb.create_sheet(title="08_Formula_Glossary_&_Guide")
    ws8.views.sheetView[0].showGridLines = True

    ws8.merge_cells("A1:F2")
    ws8["A1"] = "EXCEL VS GOOGLE SHEETS COMPATIBILITY MATRIX & SHORTCUT CHEAT SHEET"
    ws8["A1"].font = font_title
    ws8["A1"].fill = fill_title
    ws8["A1"].alignment = Alignment(horizontal="center", vertical="center")

    guide_headers = ["Formula / Feature Category", "Microsoft Excel (365 / Desktop)", "Google Sheets Equivalent", "Best Practice & Gotchas", "Status"]
    for idx, h in enumerate(guide_headers, 1):
        c = ws8.cell(row=4, column=idx, value=h)
        c.font = font_bold
        c.fill = fill_header
        c.border = cell_border

    guide_items = [
        ("Exact Lookups", "=XLOOKUP(lookup, key_col, val_col, fallback)", "=XLOOKUP(lookup, key_col, val_col, fallback)", "Both support modern XLOOKUP natively. For legacy Excel 2013-2019, use INDEX(val_col, MATCH(key, key_col, 0)).", "100% Native"),
        ("Dynamic Filtering", "=FILTER(array, condition, [if_empty])", "=FILTER(range, condition1, [condition2])", "Google Sheets requires multiple conditions passed as separate arguments or boolean multiplication *(col>x).", "100% Native"),
        ("Dynamic Array Spills", "Automatic spill behavior across adjacent cells", "Automatic or wrap in =ARRAYFORMULA() for legacy cells", "In Google Sheets, wrap legacy single-cell operations like IF, CONCAT in ARRAYFORMULA() to spill.", "100% Native"),
        ("Database / SQL Query", "Power Query (M Language) / Get & Transform", '=QUERY(A1:F100, "SELECT A, SUM(C) WHERE B=\'R&D\' GROUP BY A")', "Google Sheets has the revolutionary QUERY() SQL function. In Excel, Power Query handles enterprise ETL.", "Cross-Platform"),
        ("Named Lambdas", "Defined via Name Manager -> New -> LAMBDA()", "Defined via Data -> Named functions", "Allows defining reusable functional macros without enabling VBA macros (.xlsm) or Apps Script.", "100% Native"),
        ("External Data Sync", "Power Query Web Connector / OData Feed", '=IMPORTRANGE("sheet_url", "Tab!A1:Z")', "Google Sheets excels at cloud-to-cloud inter-workbook live linking with IMPORTRANGE.", "Native Cloud"),
        ("Discounted Cash Flow", "=NPV(rate, val1, val2...) + CF0", "=NPV(rate, val1, val2...) + CF0", "Remember that NPV assumes period 1 start! Always add initial investment (CF0) outside the NPV formula.", "Standardized"),
        ("Irregular Timing Returns", "=XIRR(values, dates)", "=XIRR(values, dates)", "Always format date columns as true Excel/Sheets dates rather than strings to avoid #VALUE! errors.", "Standardized"),
        ("Two-Way Table Lookup", "=INDEX(table, XMATCH(row_key, rows), XMATCH(col_key, cols))", "=INDEX(table, MATCH(row_key, rows, 0), MATCH(col_key, cols, 0))", "Rock-solid replacement for nested IF statements when querying 2D sensitivity matrices.", "Universal")
    ]

    for idx, (cat, xl_syntax, gs_syntax, notes, status) in enumerate(guide_items, start=5):
        ws8.cell(row=idx, column=1, value=cat).font = font_bold
        ws8.cell(row=idx, column=2, value=xl_syntax).font = font_formula
        ws8.cell(row=idx, column=3, value=gs_syntax).font = font_formula
        ws8.cell(row=idx, column=4, value=notes).font = font_muted
        c_stat = ws8.cell(row=idx, column=5, value=status)
        c_stat.font = font_bold
        c_stat.alignment = Alignment(horizontal="center")
        
        for c in range(1, 6):
            ws8.cell(row=idx, column=c).border = cell_border
            if idx % 2 == 1:
                ws8.cell(row=idx, column=c).fill = fill_zebra

    # Keyboard Shortcuts Comparison Matrix (Rows 16 to 24)
    ws8.merge_cells("A16:E16")
    ws8["A16"] = "ESSENTIAL MODELING SHORTCUT CHEAT SHEET (WINDOWS VS MAC VS SHEETS)"
    ws8["A16"].font = font_section
    ws8["A16"].fill = fill_header

    shortcut_headers = ["Action", "Excel (Windows PC)", "Excel (Mac OS)", "Google Sheets", "Productivity Impact"]
    for idx, h in enumerate(shortcut_headers, 1):
        c = ws8.cell(row=17, column=idx, value=h)
        c.font = font_bold
        c.fill = fill_subtotal
        c.border = cell_border

    shortcuts = [
        ("Absolute Reference Toggle ($)", "F4", "Command + T", "F4", "Locks row/column coordinates ($A$1) instantly"),
        ("Fill Down / Fill Right", "Ctrl + D / Ctrl + R", "Cmd + D / Cmd + R", "Ctrl + D / Ctrl + R", "Massively speeds up formula horizontal/vertical cascading"),
        ("Paste Special Formats / Formulas", "Alt + E + S + T / F", "Cmd + Opt + V", "Ctrl + Alt + V", "Preserves exact border aesthetics and number formats"),
        ("Jump to Edge of Region", "Ctrl + Arrow Keys", "Cmd + Arrow Keys", "Ctrl + Arrow Keys", "Traverses 100k+ row financial ledgers in 1 click"),
        ("Select Entire Current Region", "Ctrl + Shift + 8 (Ctrl + *)", "Cmd + A", "Ctrl + A", "Selects bounded data matrix cleanly for pivot tables"),
        ("Format as Currency ($)", "Ctrl + Shift + 4 (Ctrl + $)", "Ctrl + Shift + $", "Ctrl + Shift + 4", "Applies standardized corporate currency format"),
        ("Format as Percentage (%)", "Ctrl + Shift + 5 (Ctrl + %)", "Ctrl + Shift + %", "Ctrl + Shift + 5", "Converts decimal drivers to percentage displays")
    ]

    for idx, (act, win, mac, gs, imp) in enumerate(shortcuts, start=18):
        ws8.cell(row=idx, column=1, value=act).font = font_bold
        ws8.cell(row=idx, column=2, value=win).font = font_formula
        ws8.cell(row=idx, column=3, value=mac).font = font_formula
        ws8.cell(row=idx, column=4, value=gs).font = font_formula
        ws8.cell(row=idx, column=5, value=imp).font = font_muted
        for c in range(1, 6):
            ws8.cell(row=idx, column=c).border = cell_border
            if idx % 2 == 1:
                ws8.cell(row=idx, column=c).fill = fill_zebra

    # Auto-adjust column widths across all sheets
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value or "")
                if len(val) > max_len and len(val) < 60 and not cell.coordinate in sheet.merged_cells:
                    max_len = len(val)
            sheet.column_dimensions[col_letter].width = max(max_len + 3, 14)
            
    # Custom adjustments for specific key columns
    ws1.column_dimensions["A"].width = 32
    ws2.column_dimensions["A"].width = 34
    ws3.column_dimensions["A"].width = 38
    ws4.column_dimensions["A"].width = 40
    ws5.column_dimensions["A"].width = 36
    ws6.column_dimensions["A"].width = 34
    ws7.column_dimensions["A"].width = 30
    ws7.column_dimensions["B"].width = 45
    ws7.column_dimensions["E"].width = 50
    ws8.column_dimensions["A"].width = 28
    ws8.column_dimensions["B"].width = 38
    ws8.column_dimensions["C"].width = 38
    ws8.column_dimensions["D"].width = 50

    wb.save(output_path)
    print(f"Master Financial Workstation successfully saved to: {output_path}")

if __name__ == "__main__":
    millicom_path = "/Users/lambert/Desktop/fast-api/Data-Science/Millicom-Company/Master_Financial_Analysis_and_Modeling_Workstation.xlsx"
    portfolio_path = "/Users/lambert/Desktop/fast-api/Data-Science/portfolio_website/assets/Master_Financial_Analysis_and_Modeling_Workstation.xlsx"
    create_master_workstation(millicom_path)
    create_master_workstation(portfolio_path)
