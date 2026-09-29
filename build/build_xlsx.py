"""Builds FulfilmentLabourPlan.xlsm (live formulas plus the macro project from build/vba_project.py). All inputs are illustrative."""
import math, random, datetime as dt, json, xlsxwriter

random.seed(15)
START = dt.date(2026, 11, 2)              # a Monday
N = 56                                     # 8 weeks
WK = [1.05, 1.02, 1.00, 1.00, 1.08, 0.95, 0.70]
WEEKLVL = [0.95, 1.00, 1.15, 1.28, 1.30, 1.12, 1.00, 0.92]
BASE = 52000
demand = [round(BASE * WEEKLVL[i // 7] * WK[i % 7] * (1 + random.uniform(-.03, .03)), -1) for i in range(N)]

P = dict(ShiftHrs=9, AllowancePct=.10, IndirectPct=.05, AbsentPct=.08, WorkDays=6, Roster=230,
         OTCapHrs=2, OTTakeUp=.50, OTEff=.90, FlexFactor=.85, Wage=100, OTMult=2.0, FlexWage=120,
         ShareA=.45, ShareB=.40, ShareC=.15)
PROC = [("Stow", .90, 250), ("Pick", 1.0, 100), ("Pack", 1.0, 120), ("Ship", 1.0, 300)]
SCEN = [("Base", 1.00), ("Peak", 1.30), ("Low", .85)]

def model(mult):
    eff = (1 - P["AllowancePct"]) * (1 - P["IndirectPct"]); a = P["AbsentPct"]
    K = P["Roster"] * P["WorkDays"] / 7
    rows = []
    for u0 in demand:
        u = u0 * mult
        hrs = [u * s / uph for _, s, uph in PROC]; I = sum(hrs)
        J = I / (P["ShiftHrs"] * eff) / (1 - a)
        L = K * (1 - a) * P["ShiftHrs"] * eff
        M = max(0, I - L)
        N_ = min(M / P["OTEff"], K * (1 - a) * P["OTCapHrs"] * P["OTTakeUp"])
        O = N_ * P["OTEff"]; Pf = max(0, M - O)
        Q = math.ceil(round(Pf / (P["ShiftHrs"] * eff * P["FlexFactor"]), 9)); R = Q * P["ShiftHrs"]
        cost = K * (1 - a) * P["ShiftHrs"] * P["Wage"] + N_ * P["Wage"] * P["OTMult"] + R * P["FlexWage"]
        rows.append(dict(u=u, I=I, J=J, K=K, L=L, N=N_, Q=Q, cost=cost, util=min(I, L) / L))
    return rows

def summary(mult):
    r = model(mult); U = sum(x["u"] for x in r); C = sum(x["cost"] for x in r)
    return [U, C, C / U, sum(x["N"] for x in r), max(x["Q"] for x in r),
            sum(1 for x in r if x["Q"] > 0), sum(x["util"] for x in r) / len(r)]

if __name__ == "__main__":
    wb = xlsxwriter.Workbook("FulfilmentLabourPlan.xlsm", {"use_future_functions": True})
    wb.add_vba_project("build/vbaProject.bin")
    NAVY, GREY, BLUE, COPPER = "#22303C", "#F4EFE6", "#0000FF", "#B8642B"
    f = lambda **k: wb.add_format({"font_name": "Calibri", "font_size": 10, **k})
    title = f(bold=True, font_size=16, font_color=NAVY); hdr = f(bold=True, bg_color=NAVY, font_color="white", border=1, text_wrap=True, align="center", valign="vcenter")
    txt = f(); note = f(italic=True, font_color="#666666"); inp = f(font_color=BLUE, bg_color="#FFF9E5", border=1)
    inp_pct = f(font_color=BLUE, bg_color="#FFF9E5", border=1, num_format="0%"); inp_dec = f(font_color=BLUE, bg_color="#FFF9E5", border=1, num_format="0.00")
    calc = f(border=1, bg_color=GREY, num_format="0.000"); n0 = f(border=1, num_format="#,##0"); n1 = f(border=1, num_format="#,##0.0")
    n2 = f(border=1, num_format="#,##0.00"); dte = f(border=1, num_format="dd-mmm-yy"); pct = f(border=1, num_format="0.0%"); cell = f(border=1)
    inr = f(border=1, num_format="[$₹-4009]#,##0"); inr2 = f(border=1, num_format="[$₹-4009]#,##0.00")
    lab = f(bold=True); kpi_l = f(bold=True, bg_color=GREY, border=1)

    wd, wi, wp, ws, ww = [wb.add_worksheet(n) for n in ("Dashboard", "Inputs", "Plan", "ShiftPlan", "Weekly")]
    # ---------------- Inputs ----------------
    wi.set_column("A:A", 34); wi.set_column("B:B", 13); wi.set_column("C:C", 12); wi.set_column("D:D", 70)
    wi.write("A1", "Inputs and assumptions", title)
    wi.write("A2", "All figures are illustrative planning assumptions for a generic fulfilment centre, not data from any real site. Blue cells are inputs.", note)
    r = 3
    def put(label, name, val, unit, why, fmt):
        global r
        wi.write(r, 0, label, txt); wi.write(r, 1, val, fmt); wi.write(r, 2, unit, txt); wi.write(r, 3, why, note)
        wb.define_name(name, f"=Inputs!$B${r+1}"); r += 1
    wi.write(r, 0, "Scenario", lab); r += 1
    put("Scenario (dropdown)", "Scenario", "Base", "", "Base, Peak or Low. Scales the forecast. The scenario buttons on the Dashboard set it too.", inp)
    wi.data_validation(r - 1, 1, r - 1, 1, {"validate": "list", "source": ["Base", "Peak", "Low"]})
    wi.write(r, 0, "Demand multiplier (from table)", txt); wi.write_formula(r, 1, "=INDEX(ScenMult,MATCH(Scenario,ScenName,0))", calc)
    wb.define_name("DemandMult", f"=Inputs!$B${r+1}"); r += 2
    wi.write(r, 0, "Scenario table", lab); r += 1
    wi.write_row(r, 0, ["Scenario", "Multiplier"], hdr); r += 1
    s0 = r
    for n_, m in SCEN: wi.write(r, 0, n_, cell); wi.write(r, 1, m, inp_dec); r += 1
    wb.define_name("ScenName", f"=Inputs!$A${s0+1}:$A${r}"); wb.define_name("ScenMult", f"=Inputs!$B${s0+1}:$B${r}"); r += 1
    wi.write(r, 0, "Process standards", lab); r += 1
    wi.write_row(r, 0, ["Process", "Units per outbound unit", "UPH"], hdr); wi.write(r, 3, "UPH = units handled per productive labour hour.", note); r += 1
    for p, s, u in PROC:
        wi.write(r, 0, p, cell); wi.write(r, 1, s, inp_dec); wi.write(r, 2, u, inp)
        wb.define_name(f"Share_{p}", f"=Inputs!$B${r+1}"); wb.define_name(f"UPH_{p}", f"=Inputs!$C${r+1}"); r += 1
    wi.write(r, 3, "Stow at 0.90: inbound stow volume runs a little below outbound units.", note); r += 2
    wi.write(r, 0, "Shift and shrinkage", lab); r += 1
    put("Paid shift length", "ShiftHrs", P["ShiftHrs"], "hours", "Paid hours per associate per working day.", inp)
    put("Breaks and allowances", "AllowancePct", P["AllowancePct"], "% of shift", "Paid time not spent on standard work (breaks, PPE, handover).", inp_pct)
    put("Indirect / training time", "IndirectPct", P["IndirectPct"], "% of shift", "Meetings, training, 5S.", inp_pct)
    put("Absenteeism", "AbsentPct", P["AbsentPct"], "% of roster", "Unplanned absence; absent associates are unpaid.", inp_pct)
    wi.write(r, 0, "Productive share of paid hours", txt); wi.write_formula(r, 1, "=(1-AllowancePct)*(1-IndirectPct)", calc)
    wb.define_name("EffShare", f"=Inputs!$B${r+1}"); wi.write(r, 3, "Calculated.", note); r += 2
    wi.write(r, 0, "Permanent workforce", lab); r += 1
    put("Permanent roster", "Roster", P["Roster"], "associates", "Cross-trained pool, assumed deployable to any process.", inp)
    put("Working days per week", "WorkDays", P["WorkDays"], "days", "Weekly off is rotated; average daily availability = roster x days / 7.", inp)
    r += 1; wi.write(r, 0, "Overtime and flex labour", lab); r += 1
    put("Max overtime per associate per day", "OTCapHrs", P["OTCapHrs"], "hours", "Cap on OT hours a present associate can work.", inp)
    put("Overtime take-up", "OTTakeUp", P["OTTakeUp"], "% of present", "Share of present associates who will take OT.", inp_pct)
    put("Overtime productivity", "OTEff", P["OTEff"], "x", "Productive hours per paid OT hour.", inp_dec)
    put("Flex productivity vs permanent", "FlexFactor", P["FlexFactor"], "x", "Seasonal hires are slower in the first weeks.", inp_dec)
    r += 1; wi.write(r, 0, "Cost", lab); r += 1
    put("Permanent wage", "Wage", P["Wage"], "INR / paid hr", "Fully loaded, illustrative.", inp)
    put("Overtime multiplier", "OTMult", P["OTMult"], "x wage", "Double rate, as under the Factories Act.", inp_dec)
    put("Flex / agency wage", "FlexWage", P["FlexWage"], "INR / paid hr", "Includes agency margin.", inp)
    r += 1; wi.write(r, 0, "Shift split of attending headcount", lab); r += 1
    put("Shift A (morning)", "ShareA", P["ShareA"], "%", "Must total 100%.", inp_pct)
    put("Shift B (afternoon)", "ShareB", P["ShareB"], "%", "", inp_pct)
    put("Shift C (night)", "ShareC", P["ShareC"], "%", "", inp_pct)

    # ---------------- Plan ----------------
    cols = ["Date", "Day", "Week", "Base forecast (units)", "Scenario units", "Stow hrs", "Pick hrs", "Pack hrs", "Ship hrs",
            "Required productive hrs", "Required HC (rostered)", "Permanent HC on roster", "Permanent productive hrs", "Shortfall hrs",
            "OT paid hrs", "OT productive hrs", "Flex productive hrs needed", "Flex HC", "Flex paid hrs",
            "Regular cost", "OT cost", "Flex cost", "Total cost", "Cost per unit", "Utilisation", "Action"]
    C = {c: xlsxwriter.utility.xl_col_to_name(i) for i, c in enumerate(cols)}
    wp.write("A1", "Daily labour plan", title); wp.write("A2", "Rows are formulas driven by Inputs. Permanent labour is used first, then overtime, then flex hires.", note)
    wp.set_row(3, 42); wp.set_column(0, len(cols) - 1, 12); wp.set_column("A:A", 11); wp.set_column("Z:Z", 12)
    for i, c in enumerate(cols): wp.write(3, i, c, hdr)
    dv = f(border=1, num_format="[$₹-4009]#,##0")
    for i in range(N):
        x = 5 + i; g = lambda k: f"{C[k]}{x}"
        wp.write_datetime(x - 1, 0, dt.datetime.combine(START + dt.timedelta(i), dt.time()), dte)
        wp.write_formula(x - 1, 1, f'=TEXT({g("Date")},"ddd")', cell)
        wp.write_formula(x - 1, 2, f'=INT(({g("Date")}-$A$5)/7)+1', cell)
        wp.write(x - 1, 3, demand[i], f(border=1, num_format="#,##0", font_color=BLUE))
        wp.write_formula(x - 1, 4, f'={g("Base forecast (units)")}*DemandMult', n0)
        for j, (p, _, _) in enumerate(PROC):
            wp.write_formula(x - 1, 5 + j, f'={g("Scenario units")}*Share_{p}/UPH_{p}', n1)
        F = {
         "Required productive hrs": (f'=SUM({C["Stow hrs"]}{x}:{C["Ship hrs"]}{x})', n1),
         "Required HC (rostered)": (f'={g("Required productive hrs")}/(ShiftHrs*EffShare)/(1-AbsentPct)', n1),
         "Permanent HC on roster": ('=Roster*WorkDays/7', n1),
         "Permanent productive hrs": (f'={g("Permanent HC on roster")}*(1-AbsentPct)*ShiftHrs*EffShare', n1),
         "Shortfall hrs": (f'=MAX(0,{g("Required productive hrs")}-{g("Permanent productive hrs")})', n1),
         "OT paid hrs": (f'=MIN({g("Shortfall hrs")}/OTEff,{g("Permanent HC on roster")}*(1-AbsentPct)*OTCapHrs*OTTakeUp)', n1),
         "OT productive hrs": (f'={g("OT paid hrs")}*OTEff', n1),
         "Flex productive hrs needed": (f'=MAX(0,{g("Shortfall hrs")}-{g("OT productive hrs")})', n1),
         "Flex HC": (f'=ROUNDUP(ROUND({g("Flex productive hrs needed")}/(ShiftHrs*EffShare*FlexFactor),9),0)', n0),
         "Flex paid hrs": (f'={g("Flex HC")}*ShiftHrs', n0),
         "Regular cost": (f'={g("Permanent HC on roster")}*(1-AbsentPct)*ShiftHrs*Wage', dv),
         "OT cost": (f'={g("OT paid hrs")}*Wage*OTMult', dv),
         "Flex cost": (f'={g("Flex paid hrs")}*FlexWage', dv),
         "Total cost": (f'=SUM({g("Regular cost")}:{g("Flex cost")})', dv),
         "Cost per unit": (f'={g("Total cost")}/{g("Scenario units")}', inr2),
         "Utilisation": (f'=MIN({g("Required productive hrs")},{g("Permanent productive hrs")})/{g("Permanent productive hrs")}', pct),
         "Action": (f'=IF({g("Flex HC")}>0,"Flex hire",IF({g("OT paid hrs")}>0,"Overtime","Roster covers"))', cell)}
        for k, (fm, fmt) in F.items(): wp.write_formula(x - 1, cols.index(k), fm, fmt)
    last = 4 + N
    rng = lambda k: f"Plan!${C[k]}$5:${C[k]}${last}"
    wp.freeze_panes(4, 3)
    wp.conditional_format(f"{C['Action']}5:{C['Action']}{last}", {"type": "cell", "criteria": "==", "value": '"Flex hire"', "format": wb.add_format({"bg_color": "#F3D9C6"})})
    wp.conditional_format(f"{C['Action']}5:{C['Action']}{last}", {"type": "cell", "criteria": "==", "value": '"Overtime"', "format": wb.add_format({"bg_color": "#F6EBC8"})})

    # ---------------- ShiftPlan ----------------
    ws.write("A1", "Attending headcount by shift", title)
    ws.write("A2", "Permanent associates present (after absenteeism) plus flex hires, split by the shift shares on Inputs. Shift C is the balancing figure so rows add up.", note)
    ws.write_row(3, 0, ["Date", "Day", "Permanent present", "Flex HC", "Total attending", "Shift A", "Shift B", "Shift C"], hdr); ws.set_row(3, 30)
    ws.set_column("A:H", 13)
    for i in range(N):
        x = 5 + i
        ws.write_formula(x - 1, 0, f"=Plan!A{x}", dte); ws.write_formula(x - 1, 1, f"=Plan!B{x}", cell)
        ws.write_formula(x - 1, 2, f'=Plan!{C["Permanent HC on roster"]}{x}*(1-AbsentPct)', n1)
        ws.write_formula(x - 1, 3, f'=Plan!{C["Flex HC"]}{x}', n0)
        ws.write_formula(x - 1, 4, f"=C{x}+D{x}", n1)
        ws.write_formula(x - 1, 5, f"=ROUND(E{x}*ShareA,0)", n0); ws.write_formula(x - 1, 6, f"=ROUND(E{x}*ShareB,0)", n0)
        ws.write_formula(x - 1, 7, f"=ROUND(E{x},0)-F{x}-G{x}", n0)
    ws.freeze_panes(4, 2)

    # ---------------- Weekly ----------------
    ww.write("A1", "Weekly summary", title)
    ww.write_row(3, 0, ["Week", "Starts", "Units", "Avg required HC", "Peak required HC", "Peak flex HC", "OT paid hrs", "Flex paid hrs", "Total cost", "Cost per unit"], hdr); ww.set_row(3, 30); ww.set_column("A:J", 14)
    for k in range(N // 7):
        x = 5 + k; ww.write(x - 1, 0, k + 1, cell)
        ww.write_formula(x - 1, 1, f"=Plan!$A$5+7*(A{x}-1)", dte)
        wk = rng("Week")
        ww.write_formula(x - 1, 2, f'=SUMIFS({rng("Scenario units")},{wk},A{x})', n0)
        ww.write_formula(x - 1, 3, f'=AVERAGEIFS({rng("Required HC (rostered)")},{wk},A{x})', n1)
        ww.write_formula(x - 1, 4, f'=MAXIFS({rng("Required HC (rostered)")},{wk},A{x})', n1)
        ww.write_formula(x - 1, 5, f'=MAXIFS({rng("Flex HC")},{wk},A{x})', n0)
        ww.write_formula(x - 1, 6, f'=SUMIFS({rng("OT paid hrs")},{wk},A{x})', n0)
        ww.write_formula(x - 1, 7, f'=SUMIFS({rng("Flex paid hrs")},{wk},A{x})', n0)
        ww.write_formula(x - 1, 8, f'=SUMIFS({rng("Total cost")},{wk},A{x})', inr)
        ww.write_formula(x - 1, 9, f"=I{x}/C{x}", inr2)

    # ---------------- Dashboard ----------------
    wd.set_column("A:A", 30); wd.set_column("B:H", 15)
    wd.write("A1", "Fulfilment labour plan: 8-week view", title)
    wd.write_formula("A2", '="Scenario: "&Scenario&"  |  demand x"&TEXT(DemandMult,"0.00")&"  |  roster "&Roster&" permanent associates"', note)
    kp = [("Units shipped", f'=SUM({rng("Scenario units")})', n0, "KPI_Units"), ("Total labour cost", f'=SUM({rng("Total cost")})', inr, "KPI_Cost"),
          ("Cost per unit", "=KPI_Cost/KPI_Units", inr2, "KPI_CPU"), ("Overtime paid hours", f'=SUM({rng("OT paid hrs")})', n0, "KPI_OT"),
          ("Peak flex hires (one day)", f'=MAX({rng("Flex HC")})', n0, "KPI_FlexHC"), ("Days needing flex", f'=COUNTIF({rng("Flex HC")},">0")', n0, "KPI_FlexDays"),
          ("Average utilisation of roster", f'=AVERAGE({rng("Utilisation")})', pct, "KPI_Util")]
    for i, (l, fm, fmt, nm) in enumerate(kp):
        wd.write(3 + i, 0, l, kpi_l); wd.write_formula(3 + i, 1, fm, fmt); wb.define_name(nm, f"=Dashboard!$B${4+i}")
    wd.write("A12", "Scenario comparison", lab)
    wd.write("A13", "Refreshed by macro CompareScenarios. Values shown were calculated with the delivered inputs; re-run after changing Inputs.", note)
    wd.write_row(13, 0, ["Scenario", "Units", "Total cost", "Cost per unit", "OT paid hrs", "Peak flex HC", "Days needing flex", "Avg utilisation"], hdr); wd.set_row(13, 30)
    fm2 = [None, n0, inr, inr2, n0, n0, n0, pct]
    for i, (n_, m) in enumerate(SCEN):
        s = summary(m); wd.write(14 + i, 0, n_, cell)
        for j, v in enumerate(s): wd.write(14 + i, 1 + j, v, fm2[1 + j])
    wb.define_name("CompareTop", "=Dashboard!$A$15")
    ch = wb.add_chart({"type": "line"})
    ds = f"Plan!${C['Date']}$5:${C['Date']}${last}"
    ch.add_series({"name": "Required HC", "categories": ds, "values": f"=Plan!${C['Required HC (rostered)']}$5:${C['Required HC (rostered)']}${last}", "line": {"color": COPPER, "width": 2.5}})
    ch.add_series({"name": "Permanent HC on roster", "categories": ds, "values": f"=Plan!${C['Permanent HC on roster']}$5:${C['Permanent HC on roster']}${last}", "line": {"color": NAVY, "width": 2.25, "dash_type": "dash"}})
    ch.set_title({"name": "Required vs permanent headcount", "name_font": {"size": 11}}); ch.set_legend({"position": "bottom"})
    ch.set_x_axis({"date_axis": True, "num_format": "dd-mmm", "major_gridlines": {"visible": False}}); ch.set_y_axis({"name": "Headcount", "num_format": "0", "major_gridlines": {"visible": True, "line": {"color": "#E3DDD0"}}})
    ch.set_size({"width": 760, "height": 300}); wd.insert_chart("A20", ch)
    c2 = wb.add_chart({"type": "column"})
    c2.add_series({"name": "Flex HC", "categories": ds, "values": f"=Plan!${C['Flex HC']}$5:${C['Flex HC']}${last}", "fill": {"color": COPPER}, "gap": 40})
    c2.set_title({"name": "Flex hires needed per day", "name_font": {"size": 11}}); c2.set_legend({"none": True})
    c2.set_x_axis({"date_axis": True, "num_format": "dd-mmm"}); c2.set_y_axis({"major_gridlines": {"visible": True, "line": {"color": "#E3DDD0"}}}); c2.set_size({"width": 760, "height": 250}); wd.insert_chart("A36", c2)
    for i, (cap, mac) in enumerate([("Base scenario", "RunScenarioBase"), ("Peak scenario", "RunScenarioPeak"), ("Low scenario", "RunScenarioLow"),
                                    ("Compare scenarios", "CompareScenarios"), ("Check inputs", "ValidateInputs"), ("Export PDF", "ExportDashboardPDF")]):
        wd.insert_button(3 + i, 3, {"macro": mac, "caption": cap, "width": 130, "height": 20, "y_offset": 1})
    for sh, col in ((wd, COPPER), (wi, "#C9A227"), (wp, NAVY), (ws, NAVY), (ww, NAVY)):
        sh.set_tab_color(col); sh.hide_gridlines(2); sh.set_landscape(); sh.fit_to_pages(1, 0); sh.set_paper(9)
    wp.repeat_rows(3); ws.repeat_rows(3)
    wd.set_first_sheet(); wd.activate()
    wb.close()
    json.dump({n_: summary(m) for n_, m in SCEN}, open("build/expected.json", "w"))
