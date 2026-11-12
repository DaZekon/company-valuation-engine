"""Formula-driven native Excel export using openpyxl; no external workbook links."""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.properties import CalcProperties

from .model import Drivers, Valuation

NAVY, TEAL, BLUE, GREEN, BLACK = "17324D", "286F78", "0000FF", "008000", "000000"
NUM = '#,##0.0;(#,##0.0);"–"'
PCT = '0.0%;(0.0%);"–"'
PRICE = '0.00;(0.00);"–"'
GROUPS = {"growth":22, "margin":27, "da_ratio":32, "capex_ratio":37, "nwc_ratio":42}
LABELS = {"growth":"Revenue growth", "margin":"EBIT margin", "da_ratio":"D&A / revenue", "capex_ratio":"Capex / revenue", "nwc_ratio":"Operating NWC / revenue"}
HIST = {8:"revenue",10:"ebit",12:"net_income",13:"cfo",14:"capex",16:"da",17:"operating_nwc",20:"cash",21:"debt",22:"diluted_weighted_avg_shares"}


def write(ws, cell: str, value, fmt: str | None = None, kind: str = "formula"):
    c = ws[cell]
    c.value = value
    c.font = Font(name="Arial", size=10, color={"input":BLUE,"source":GREEN,"formula":BLACK,"text":NAVY}[kind])
    c.alignment = Alignment(vertical="center", horizontal="right" if isinstance(value,(int,float)) or (isinstance(value,str) and value.startswith("=")) else "left")
    if fmt:
        c.number_format = fmt
    if kind == "input":
        c.fill = PatternFill("solid", fgColor="FFF4CF")
    return c


def section(ws, row: int, title: str, end: int = 8):
    for col in range(2,end+1):
        c=ws.cell(row,col)
        c.fill=PatternFill("solid",fgColor=NAVY)
        c.font=Font(name="Arial",size=10,color="FFFFFF",bold=True)
    ws.cell(row,2,title)


def base_sheet(wb, name: str, title: str):
    ws=wb.create_sheet(name)
    ws.sheet_view.showGridLines=False
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.sheet_properties.outlinePr.summaryRight=False
    ws.sheet_properties.tabColor=NAVY if name=="Summary" else TEAL if name=="Inputs" else "C9D5DE"
    ws.column_dimensions["A"].width=3
    ws.column_dimensions["B"].width=38
    ws.column_dimensions["C"].width=19
    for col in range(4,12): ws.column_dimensions[letter(col)].width=17
    for row in range(1,130): ws.row_dimensions[row].height=18
    write(ws,"B2",title,kind="text").font=Font(name="Arial",size=16,color=NAVY,bold=True)
    ws.freeze_panes="D7" if name not in ("Summary","DCF") else None
    ws.page_setup.orientation="landscape"
    ws.page_setup.paperSize=ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth=1
    ws.page_setup.fitToHeight=1
    ws.print_options.horizontalCentered=True
    ws.print_area="B1:K40"
    return ws


def note(ws, row: int, text: str, end: str="H", height: int=32):
    ws.merge_cells(f"B{row}:{end}{row}")
    write(ws,f"B{row}",text,kind="text").alignment=Alignment(wrap_text=True,vertical="center")
    ws.row_dimensions[row].height=height


def decimal_validation(ws, cells: str, lower: float, upper: float):
    dv=DataValidation(type="decimal",operator="between",formula1=str(lower),formula2=str(upper),allow_blank=False)
    dv.showErrorMessage=True
    dv.errorTitle="Invalid assumption"
    dv.error=f"Enter a number between {lower} and {upper}."
    dv.errorStyle="stop"
    ws.add_data_validation(dv)
    dv.add(cells)


def guarded(expr: str, *refs: str) -> str:
    return f'=IF(COUNT({",".join(refs)})={len(refs)},{expr},"n.a.")'


def forecast_block(ws, start: int, assumption_rows: dict, source_sheet: str, last_hist: str, active: bool=True):
    """Visible five-year operating forecast. Same rows power the selected case and comparison cases."""
    # Relative rows: year0, growth3, revenue4, margin5, EBIT6, taxrate7, taxes8,
    # NOPAT9, DA ratio10, DA11, capex ratio12, capex13, NWC ratio14, NWC15, dNWC16, FCFF17.
    labels={0:"Forecast year",1:"Forecast period end",3:"Revenue growth",4:"Revenue",5:"EBIT margin",6:"EBIT",7:"Operating tax rate",8:"Operating cash taxes",9:"NOPAT",10:"D&A / revenue",11:"D&A",12:"Capex / revenue",13:"Capex",14:"Operating NWC / revenue",15:"Operating NWC",16:"Increase in operating NWC",17:"Unlevered free cash flow (FCFF)"}
    for off,label in labels.items(): write(ws,f"B{start+off}",label,kind="text")
    for i,col in enumerate("DEFGH"):
        def ref(off): return f"{col}{start+off}"
        prev=letter(3+i)
        write(ws,ref(0),i+1,"0")
        write(ws,ref(1),f"=EDATE('Inputs'!$C$6,{ref(0)}*12)","yyyy-mm-dd")
        for off,name in ((3,"growth"),(5,"margin"),(10,"da_ratio"),(12,"capex_ratio"),(14,"nwc_ratio")):
            source=f"'{source_sheet}'!{col}{assumption_rows[name]}"
            bounds=f"{source}>-1" if name=="growth" else f"AND({source}>=-1,{source}<=1)" if name in ("margin","nwc_ratio") else f"{source}>=0"
            write(ws,ref(off),f'=IF(COUNT({source})=1,IF({bounds},{source},"n.a."),"n.a.")',PCT)
        revprev=f"'Historical'!{last_hist}8" if i==0 else f"{prev}{start+4}"
        nwcprev=f"'Historical'!{last_hist}17" if i==0 else f"{prev}{start+15}"
        expressions={4:(f"{revprev}*(1+{ref(3)})",(revprev,ref(3))),
                     6:(f"{ref(4)}*{ref(5)}",(ref(4),ref(5))),
                     7:("IF(AND('Inputs'!$C$15>=0,'Inputs'!$C$15<=1),'Inputs'!$C$15,\"n.a.\")",("'Inputs'!$C$15",)),
                     8:(f"MAX({ref(6)},0)*{ref(7)}",(ref(6),ref(7))),
                     9:(f"{ref(6)}-{ref(8)}",(ref(6),ref(8))),
                     11:(f"{ref(4)}*{ref(10)}",(ref(4),ref(10))),
                     13:(f"{ref(4)}*{ref(12)}",(ref(4),ref(12))),
                     15:(f"{ref(4)}*{ref(14)}",(ref(4),ref(14))),
                     16:(f"{ref(15)}-{nwcprev}",(ref(15),nwcprev)),
                     17:(f"{ref(9)}+{ref(11)}-{ref(13)}-{ref(16)}",(ref(9),ref(11),ref(13),ref(16)))}
        for off,(expr,refs) in expressions.items(): write(ws,ref(off),guarded(expr,*refs),PCT if off==7 else NUM)
    for c in ws[start+17][1:8]:
        c.border=Border(top=Side(style="thin",color=NAVY))
        c.font=Font(name="Arial",size=10,bold=True,color=BLACK)


def stable_terminal(ws, start: int, growth_ref: str, dest_col: str="J"):
    """Year 6 uses terminal growth and year 5 operating ratios."""
    def d(off): return f"{dest_col}{start+off}"
    def h(off): return f"H{start+off}"
    write(ws,d(0),"Stable year 6",kind="text")
    for off,formula,refs in [
        (3,growth_ref,(growth_ref,)),(4,f"{h(4)}*(1+{d(3)})",(h(4),d(3))),
        (5,h(5),(h(5),)),(6,f"{d(4)}*{d(5)}",(d(4),d(5))),
        (7,h(7),(h(7),)),(8,f"MAX({d(6)},0)*{d(7)}",(d(6),d(7))),
        (9,f"{d(6)}-{d(8)}",(d(6),d(8))),
        (10,h(10),(h(10),)),(11,f"{d(4)}*{d(10)}",(d(4),d(10))),
        (12,h(12),(h(12),)),(13,f"{d(4)}*{d(12)}",(d(4),d(12))),
        (14,h(14),(h(14),)),(15,f"{d(4)}*{d(14)}",(d(4),d(14))),
        (16,f"{d(15)}-{h(15)}",(d(15),h(15))),
        (17,f"{d(9)}+{d(11)}-{d(13)}-{d(16)}",(d(9),d(11),d(13),d(16)))]:
        write(ws,d(off),guarded(formula,*refs),PCT if off in (3,5,7,10,12,14) else NUM)


def make_workbook(history: pd.DataFrame, config: dict, results: dict[str, tuple[pd.DataFrame, Valuation]],
                  path: str | Path, provenance: list[dict] | None=None) -> Path:
    wb=Workbook()
    wb.remove(wb.active)
    wb.calculation=CalcProperties(calcId=191029,fullCalcOnLoad=True,forceFullCalc=True,calcMode="auto")
    for name,title in [("Summary","Company valuation"),("Inputs","Financial inputs and assumptions"),("Historical","Historical financials"),("Projection","Five-year operating forecast"),("DCF","Discounted cash flow valuation"),("Sensitivity","Sensitivity and scenarios"),("Sources","Sources and limitations")]:
        base_sheet(wb,name,title)
    inp=wb["Inputs"]
    for r,label,val in [(4,"Company",config["company"]),(5,"Ticker",config["ticker"]),(6,"Valuation date",date.fromisoformat(config["valuation_date"])),(7,"Selected case (1 / 2 / 3)",2)]:
        write(inp,f"B{r}",label,kind="text");write(inp,f"C{r}",val,"yyyy-mm-dd" if r==6 else None,kind="input" if r==7 else "source")
    write(inp,"D7","1 Bear   2 Base   3 Bull",kind="text")
    decimal_validation(inp,"C7",1,3)
    inp.data_validations.dataValidation[-1].type="whole"
    write(inp,"C8",'=IF(ISNUMBER(C7),IF(AND(C7=INT(C7),C7>=1,C7<=3),CHOOSE(C7,"Bear","Base","Bull"),"INVALID CASE"),"INVALID CASE")')
    capital=config["capital"]
    for r,name,label in [(10,"risk_free_rate","Risk-free rate"),(11,"beta","Levered equity beta"),(12,"equity_risk_premium","Equity risk premium"),(13,"cost_of_debt","Pre-tax cost of debt"),(14,"debt_weight","Market-value debt weight"),(15,"tax_rate","Operating tax rate")]:
        write(inp,f"B{r}",label,kind="text");write(inp,f"C{r}",capital[name],"0.00x" if r==11 else PCT,"input")
        decimal_validation(inp,f"C{r}",-5 if r==11 else 0,5 if r==11 else 1)
    for r,label,formula in [(16,"CAPM cost of equity",guarded("C10+C11*C12","C10","C11","C12")),(17,"Calculated WACC",guarded("C16*(1-C14)+C13*(1-C15)*C14","C16","C14","C13","C15")),(18,"Active WACC","=INDEX(E50:E52,C7)"),(19,"Active terminal growth","=INDEX(F50:F52,C7)")]:
        if r in (18,19):formula=f'=IF(C8="INVALID CASE","n.a.",{formula[1:]})'
        write(inp,f"B{r}",label,kind="text");write(inp,f"C{r}",formula,PCT)
    for i,col in enumerate("DEFGH"): write(inp,f"{col}21",i+1,"0",kind="text")
    for name,top in GROUPS.items():
        write(inp,f"B{top}",LABELS[name]+" (active)",kind="text")
        for caseidx,case in enumerate(("Bear","Base","Bull"),1):
            r=top+caseidx
            write(inp,f"B{r}",case,kind="text")
            for i,col in enumerate("DEFGH"): write(inp,f"{col}{r}",config["scenarios"][case][name][i],PCT,"input")
            decimal_validation(inp,f"D{r}:H{r}",-.9999 if name=="growth" else -1 if name in ("margin","nwc_ratio") else 0,10 if name=="growth" else 1)
        for col in "DEFGH":
            chosen=f"INDEX({col}{top+1}:{col}{top+3},$C$7)"
            write(inp,f"{col}{top}",f'=IF($C$8="INVALID CASE","n.a.",IF(COUNT({chosen})=1,{chosen},"n.a."))',PCT)
    section(inp,48,"Scenario discount assumptions",6)
    for col,text in zip("BCDEF",("Scenario","WACC spread","WACC override","Effective WACC","Terminal growth")): write(inp,f"{col}49",text,kind="text")
    for r,case in zip((50,51,52),("Bear","Base","Bull")):
        s=config["scenarios"][case]
        for col,val,kind in [("B",case,"text"),("C",s.get("wacc_spread",0),"input"),("D",s.get("wacc_override"),"input"),("F",s["terminal_growth"],"input")]:write(inp,f"{col}{r}",val,PCT if col!="B" else None,kind)
        write(inp,f"E{r}",f'=IF(D{r}="",IF(COUNT($C$17,C{r})=2,$C$17+C{r},"n.a."),IF(ISNUMBER(D{r}),D{r},"n.a."))',PCT)
    section(inp,54,"Reviewed enterprise-to-equity adjustments",8)
    bridge=config["bridge"]
    for r,label,val in [(55,"Excess cash added",bridge["cash_adjustment"]),(56,"Debt deducted",bridge["debt_adjustment"]),(57,"Other adjustments (signed)",bridge["other_adjustments"]),(58,"Point-in-time diluted shares (m)",config.get("diluted_shares"))]:
        write(inp,f"B{r}",label,kind="text");write(inp,f"C{r}",val,NUM,"input")
    decimal_validation(inp,"C55:C56",0,1e12)
    decimal_validation(inp,"C58",.000001,1e12)
    note(inp,60,"Blank WACC override uses calculated WACC plus the scenario spread. All amounts and shares are in millions.")
    note(inp,62,"Blue = editable input. Black = formula. Green = source figure. Fictional sample figures are illustrative inputs.")
    note(inp,64,"Bridge inputs require review. Use excess cash, market-value debt when available, and subtract preferred stock/NCI/other claims via signed adjustments.",height=40)
    inp.print_area="B1:J65";inp.page_setup.fitToHeight=2

    hist=wb["Historical"]
    note(hist,3,"Annual USD millions; diluted weighted-average shares in millions. Blank source values mean unavailable.")
    labels={5:"Fiscal start",8:"Revenue",9:"Revenue growth",10:"Operating income / EBIT proxy",11:"Operating margin",12:"Net income",13:"Cash flow from operations",14:"Capex (positive outflow)",15:"CFO less capex (not FCFF)",16:"D&A (review scope)",17:"Operating NWC",18:"Increase in operating NWC",20:"Cash and equivalents",21:"Debt (review definition)",22:"Diluted weighted-average EPS shares"}
    for r,label in labels.items():write(hist,f"B{r}",label,kind="text")
    sourcekind="input" if config.get("illustrative") else "source"
    for i,row in history.iterrows():
        col=letter(i+4); prev=letter(i+3)
        write(hist,f"{col}4",row.fiscal_end.to_pydatetime(),"yyyy-mm-dd","source")
        write(hist,f"{col}5",row.fiscal_start.to_pydatetime(),"yyyy-mm-dd","source")
        for r,field in HIST.items():write(hist,f"{col}{r}",None if pd.isna(row[field]) else float(row[field]),NUM,sourcekind)
        if i:write(hist,f"{col}9",f'=IF(COUNT({col}8,{prev}8)=2,IF({prev}8=0,"n.a.",{col}8/{prev}8-1),"n.a.")',PCT)
        write(hist,f"{col}11",f'=IF(COUNT({col}10,{col}8)=2,IF({col}8=0,"n.a.",{col}10/{col}8),"n.a.")',PCT)
        write(hist,f"{col}15",guarded(f"{col}13-{col}14",f"{col}13",f"{col}14"),NUM)
        if i:write(hist,f"{col}18",guarded(f"{col}17-{prev}17",f"{col}17",f"{prev}17"),NUM)
    last=letter(len(history)+3)
    note(hist,25,"Operating NWC excludes cash, debt and financing balances. Historical EPS share averages are not a current capital-structure denominator.",height=38)
    note(hist,27,"CFO less capex is a reported cash-flow metric. It is neither FCFF nor complete FCFE without financing and other adjustments.",height=38)

    proj=wb["Projection"]
    write(proj,"C3","='Inputs'!C8")
    forecast_block(proj,4,GROUPS,"Inputs",last)
    stable_terminal(proj,4,"'Inputs'!$C$19")
    for col in "DEFGH":
        write(proj,f"{col}23",f"='Inputs'!$C$18",PCT)
        write(proj,f"{col}24",f'=IF(ISNUMBER({col}23),IF(AND({col}23>0,{col}23<=1),1/(1+{col}23)^{col}4,"n.a."),"n.a.")',"0.000x")
        write(proj,f"{col}25",guarded(f"{col}21*{col}24",f"{col}21",f"{col}24"),NUM)
    for r,l in ((23,"Active WACC"),(24,"End-year discount factor"),(25,"Present value of annual FCFF")):write(proj,f"B{r}",l,kind="text")
    note(proj,28,"Cash taxes = MAX(EBIT, 0) × tax rate. Losses receive no immediate cash tax benefit. NOL carryforwards are not modeled.",end="K",height=34)
    note(proj,30,"FCFF = NOPAT + D&A − capex − increase in operating NWC. Year 6 recalculates NWC investment at terminal growth, using year 5 margins and investment ratios.",end="K",height=40)

    dcf=wb["DCF"]
    write(dcf,"C3","='Inputs'!C8")
    fields={6:("WACC","='Inputs'!C18",PCT),7:("Terminal growth","='Inputs'!C19",PCT),8:("Year 6 FCFF","='Projection'!J21",NUM)}
    for r,(label,formula,fmt) in fields.items():write(dcf,f"B{r}",label,kind="text");write(dcf,f"C{r}",formula,fmt)
    # Guards stay with the valuation they protect. Audit outputs never drive the model.
    write(dcf,"B10","Perpetual-growth status",kind="text")
    write(dcf,"C10",'=IF(COUNT(C6:C8)<>3,"MISSING INPUT",IF(AND(C6>0,C6<=1,C7>-1,C7<C6,C8>0,COUNT(\'Projection\'!D21:H21)=5,COUNT(\'Inputs\'!C55:C57)=3),"VALID","INVALID ASSUMPTIONS"))')
    formulas={12:("PV of years 1–5 FCFF",'=IF(C10="VALID",SUM(\'Projection\'!D25:H25),"n.a.")'),
              13:("Terminal value at year 5",'=IF(C10="VALID",C8/(C6-C7),"n.a.")'),
              14:("PV of terminal value",'=IF(C10="VALID",C13/(1+C6)^5,"n.a.")'),
              16:("Enterprise value",guarded("SUM(C12,C14)","C12","C14")),
              18:("Add: reviewed excess cash",guarded("'Inputs'!C55","'Inputs'!C55")),19:("Less: reviewed debt",guarded("'Inputs'!C56","'Inputs'!C56")),20:("Other adjustments (signed)",guarded("'Inputs'!C57","'Inputs'!C57")),
              22:("Equity value",guarded("C16+C18-C19+C20","C16","C18","C19","C20")),
              24:("Point-in-time diluted shares (m)",'=IF(ISNUMBER(\'Inputs\'!C58),IF(\'Inputs\'!C58>0,\'Inputs\'!C58,"n.a."),"n.a.")'),
              25:("Implied value per share",guarded("C22/C24","C22","C24")),
              27:("PV terminal / enterprise value",'=IF(COUNT(C14,C16)=2,IF(C16=0,"n.a.",C14/C16),"n.a.")')}
    for r,(label,formula) in formulas.items():write(dcf,f"B{r}",label,kind="text");write(dcf,f"C{r}",formula,PCT if r==27 else PRICE if r==25 else NUM)
    for r in (16,22,25):
        dcf[f"C{r}"].font=Font(name="Arial",size=14,color=NAVY,bold=True)
        dcf[f"C{r}"].border=Border(top=Side(style="thin",color=NAVY))
    note(dcf,30,"Retrospective valuation at the last fiscal year-end. Forecasts are discounted at year-end. No stub-period or mid-year discounting.")
    note(dcf,32,"A positive, sustainable terminal FCFF and WACC above terminal growth are required. Negative explicit-period FCFF remains in the valuation.")

    sen=wb["Sensitivity"]
    note(sen,3,"Enterprise value in USD millions. Sensitivity varies WACC and terminal growth; operating assumptions stay at the selected case.",end="J")
    write(sen,"B4","Selected case",kind="text");write(sen,"C4","='Inputs'!C8")
    section(sen,5,"WACC versus terminal growth (enterprise value)",8)
    for j,offset in enumerate((-.01,-.005,0,.005,.01),4):write(sen,f"{letter(j)}6",f"='Inputs'!C19+({offset})",PCT)
    for i,offset in enumerate((-.02,-.01,0,.01,.02),7):
        write(sen,f"C{i}",f"='Inputs'!C18+({offset})",PCT)
        for j in range(4,9):
            col=letter(j)
            # Simple explicit discounting plus a visibly defined stable year-6 build.
            pv="+".join(f"'Projection'!{c}21/(1+$C{i})^{n}" for n,c in enumerate("DEFGH",1))
            rev6=f"'Projection'!$H$8*(1+{col}$6)"
            terminal=f"(({rev6})*('Projection'!$H$9*(1-'Inputs'!$C$15)+'Projection'!$H$14-'Projection'!$H$16)-(({rev6})*'Projection'!$H$18-'Projection'!$H$19))"
            valid=f"AND($C{i}>0,$C{i}<=1,{col}$6>-1,{col}$6<$C{i},{terminal}>0)"
            # For negative EBIT, taxes are zero rather than a cash refund.
            terminal=f"(({rev6})*'Projection'!$H$9-MAX(({rev6})*'Projection'!$H$9,0)*'Inputs'!$C$15+({rev6})*('Projection'!$H$14-'Projection'!$H$16)-(({rev6})*'Projection'!$H$18-'Projection'!$H$19))"
            valid=f"AND($C{i}>0,$C{i}<=1,{col}$6>-1,{col}$6<$C{i},{terminal}>0)"
            write(sen,f"{col}{i}",f'=IF(\'DCF\'!$C$10<>"VALID","n.a.",IF({valid},{pv}+({terminal})/($C{i}-{col}$6)/(1+$C{i})^5,"n.a."))',NUM)
    sen.conditional_formatting.add("D7:H11",ColorScaleRule(start_type="min",start_color="F5D4CB",mid_type="percentile",mid_value=50,mid_color="FFFFFF",end_type="max",end_color="B8DAD1"))
    section(sen,14,"Scenario comparison (independent formulas)",8)
    for col,case,ix in zip("DEF",("Bear","Base","Bull"),(0,1,2)):
        top=25+ix*26
        write(sen,f"{col}15",case,kind="text")
        for r,label in ((16,"Enterprise value"),(17,"Equity value"),(18,"Implied value per share"),(19,"WACC"),(20,"Terminal growth")):
            write(sen,f"B{r}",label,kind="text")
            target={16:f"J{top+20}",17:f"J{top+21}",18:f"J{top+22}",19:f"'Inputs'!E{50+ix}",20:f"'Inputs'!F{50+ix}"}[r]
            write(sen,f"{col}{r}",f"={target}",PCT if r>=19 else PRICE if r==18 else NUM)
        section(sen,top-1,case+" operating forecast",10)
        forecast_block(sen,top,{n:r+ix+1 for n,r in GROUPS.items()},"Inputs",last,False)
        stable_terminal(sen,top,f"'Inputs'!F{50+ix}")
        w=f"'Inputs'!E{50+ix}";g=f"'Inputs'!F{50+ix}";terminal=f"J{top+17}"
        write(sen,f"I{top+19}","Case status",kind="text")
        write(sen,f"J{top+19}",f'=IF(COUNT({w},{g},{terminal})<>3,"MISSING INPUT",IF(AND({w}>0,{w}<=1,{g}>-1,{g}<{w},{terminal}>0,COUNT(D{top+17}:H{top+17})=5),"VALID","INVALID"))')
        pv="+".join(f"{c}{top+17}/(1+{w})^{n}" for n,c in enumerate("DEFGH",1))
        for off,label,formula in [(20,"Enterprise value",f'=IF(J{top+19}="VALID",{pv}+{terminal}/({w}-{g})/(1+{w})^5,"n.a.")'),
                                  (21,"Equity value",guarded(f"J{top+20}+'Inputs'!$C$55-'Inputs'!$C$56+'Inputs'!$C$57",f"J{top+20}","'Inputs'!$C$55","'Inputs'!$C$56","'Inputs'!$C$57")),
                                  (22,"Value per share",f'=IF(COUNT(J{top+21},\'Inputs\'!$C$58)=2,IF(\'Inputs\'!$C$58>0,J{top+21}/\'Inputs\'!$C$58,"n.a."),"n.a.")')]:
            write(sen,f"I{top+off}",label,kind="text");write(sen,f"J{top+off}",formula,PRICE if off==22 else NUM)
    sen.print_area="B1:K100";sen.page_setup.fitToHeight=3

    summ=wb["Summary"]
    note(summ,4,config["company"],end="J",height=28)
    note(summ,6,"ILLUSTRATIVE TRAINING DATA — fictional company, hypothetical assumptions, no real security." if config.get("illustrative") else "Reviewed financial inputs. Assumption-driven valuation as of the stated fiscal year-end.",end="J",height=32)
    for r,label,formula,fmt in [(8,"Valuation date","='Inputs'!C6","yyyy-mm-dd"),(9,"Selected scenario","='Inputs'!C8",None),(10,"Model status","='DCF'!C10",None),(12,"Enterprise value (USD m)","='DCF'!C16",NUM),(13,"Equity value (USD m)","='DCF'!C22",NUM),(14,"Implied value per share (USD)","='DCF'!C25",PRICE),(16,"WACC","='DCF'!C6",PCT),(17,"Terminal growth","='DCF'!C7",PCT),(18,"PV terminal / enterprise value","='DCF'!C27",PCT)]:
        write(summ,f"B{r}",label,kind="text");write(summ,f"C{r}",formula,fmt)
    chart=BarChart();chart.title="Projected unlevered free cash flow";chart.y_axis.title="USD millions";chart.x_axis.title="Forecast year";chart.legend=None
    chart.add_data(Reference(proj,min_col=4,max_col=8,min_row=21,max_row=21),from_rows=True)
    chart.set_categories(Reference(proj,min_col=4,max_col=8,min_row=4,max_row=4));chart.width=19;chart.height=9
    summ.add_chart(chart,"E8")
    line=LineChart();line.title="Revenue forecast";line.y_axis.title="USD millions";line.legend=None
    line.add_data(Reference(proj,min_col=4,max_col=8,min_row=8,max_row=8),from_rows=True);line.set_categories(Reference(proj,min_col=4,max_col=8,min_row=4,max_row=4));line.width=19;line.height=8
    summ.add_chart(line,"E25")
    section(summ,23,"How to use the model",3)
    for row,text in [(25,"1. Review historical inputs and Sources."),(27,"2. Edit blue assumptions on Inputs."),(29,"3. Change the case selector: 1 / 2 / 3."),(31,"4. Review DCF and sensitivity results."),(34,"Excel recalculates formulas on opening."),(36,"Read the methodology and limitations.")]:
        summ.merge_cells(start_row=row,start_column=2,end_row=row,end_column=3)
        write(summ,f"B{row}",text,kind="text")
    summ.print_area="B1:K41"

    src=wb["Sources"]
    source_notes=["Units: USD millions and shares in millions after normalization. No FX conversion.",
                  f"Valuation date: {config['valuation_date']}. Retrospective fiscal year-end model; not a current market price target.",
                  "Method: FCFF = EBIT − MAX(EBIT,0) × tax + D&A − capex − change in operating NWC.",
                  "Cost of equity = risk-free rate + beta × equity risk premium. WACC = Ke × (1 − debt weight) + Kd × (1 − tax) × debt weight.",
                  "Terminal value = stable year-6 FCFF / (WACC − g). Year-6 margins and investment ratios equal year 5; NWC investment reflects g.",
                  "A positive terminal FCFF does not prove the implied reinvestment rate or long-run ROIC is economically sustainable. Review both.",
                  "D&A and capex ratios are simple operating assumptions. No full PP&E schedule, stock compensation dilution schedule, NOLs or balance-sheet model.",
                  "Other claims, leases, minority interests, nonoperating assets and share dilution require explicit bridge review. Zero is an assumption, not missing data.",
                  config["review_note"],"Assumptions: "+config["assumption_source"]+" Dated "+config["assumption_date"],
                  "Equity bridge: "+config["bridge"]["source"],"Shares: "+config.get("shares_source","Unavailable. Per-share valuation suppressed."),
                  "SEC API docs: https://www.sec.gov/search-filings/edgar-application-programming-interfaces",
                  "SEC access guidance: https://www.sec.gov/about/developer-resources",
                  "FCFF method: https://www.cfainstitute.org/insights/professional-learning/refresher-readings/2026/free-cash-flow-valuation",
                  "Sensitivity uses explicit formulas, not Excel What-If Data Tables. Case comparisons have visible forecasts below the matrix."]
    for i,text in enumerate(source_notes,4):note(src,i,text,end="K",height=34 if i not in (7,11,14) else 42)
    r=22
    section(src,r,"Historical input provenance",11)
    headings=("Field","Fiscal start","Fiscal end","Tag / source","Raw value","Unit","Filed","Accession","Retrieved","URL")
    for col,title in enumerate(headings,2):write(src,f"{letter(col)}{r+1}",title,kind="text")
    records=provenance or [dict(field="All illustrative fields" if config.get("illustrative") else "CSV inputs (review file)",fiscal_start=str(row.fiscal_start.date()),fiscal_end=str(row.fiscal_end.date()),tag=row.source,raw_value="See Historical",unit="USD m / shares m",filed="n.a.",accession="n.a.",retrieved_at=str(row.retrieved_at),url=row.source) for _,row in history.iterrows()]
    if provenance:
        records=records+[dict(field="Reviewed CSV row",fiscal_start=str(row.fiscal_start.date()),fiscal_end=str(row.fiscal_end.date()),tag=row.source,raw_value="See Historical",unit="USD m / shares m",retrieved_at=str(row.retrieved_at),url=row.source) for _,row in history.iterrows()]
    keys=("field","fiscal_start","fiscal_end","tag","raw_value","unit","filed","accession","retrieved_at","url")
    for rowidx,record in enumerate(records,r+2):
        src.row_dimensions[rowidx].height=62
        for col,key in enumerate(keys,2):
            cell=write(src,f"{letter(col)}{rowidx}",record.get(key,""),kind="source")
            cell.alignment=Alignment(wrap_text=True,vertical="top")
            if key=="url" and str(record.get(key,"")).startswith("https://"):cell.hyperlink=record[key]
    for col,width in {"B":30,"C":16,"D":16,"E":44,"F":18,"G":20,"H":16,"I":28,"J":30,"K":65}.items():src.column_dimensions[col].width=width
    src.print_area=f"B1:K{max(30,24+len(records))}"
    for ws in wb:
        for row in ws:
            for cell in row:
                if cell.data_type=="f":cell.font=Font(name="Arial",size=10,color=BLACK,bold=cell.row in (16,22,25) and ws.title=="DCF")
        ws.sheet_view.zoomScale=85
    Path(path).parent.mkdir(parents=True,exist_ok=True)
    wb.save(path)
    return Path(path)
