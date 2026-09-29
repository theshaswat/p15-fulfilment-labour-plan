Attribute VB_Name = "modLabourPlan"
Option Explicit

' Fulfilment labour plan macros. The Plan sheet is formula-driven;
' these routines switch scenarios, check inputs, compare scenarios and export the dashboard.

Private Function Nm(ByVal n As String) As Range
    Set Nm = ThisWorkbook.Names(n).RefersToRange
End Function

Public Sub RunScenarioBase()
    SetScenario "Base"
End Sub

Public Sub RunScenarioPeak()
    SetScenario "Peak"
End Sub

Public Sub RunScenarioLow()
    SetScenario "Low"
End Sub

Private Sub SetScenario(ByVal s As String)
    Nm("Scenario").Value = s
    Application.Calculate
    ThisWorkbook.Worksheets("Dashboard").Activate
End Sub

' Returns True when every input is inside a sensible range. Shows the problems unless quiet is True.
Public Function ValidateInputs(Optional ByVal quiet As Boolean = False) As Boolean
    Dim msg As String, c As Range, p As Variant
    For Each p In Array("ShiftHrs", "Roster", "OTMult", "Wage", "FlexWage", "OTEff", "FlexFactor")
        If Not IsNumeric(Nm(CStr(p)).Value) Then
            msg = msg & "- " & p & " must be a number" & vbCrLf
        ElseIf Nm(CStr(p)).Value <= 0 Then
            msg = msg & "- " & p & " must be above zero" & vbCrLf
        End If
    Next p
    If Nm("ShiftHrs").Value > 16 Then msg = msg & "- ShiftHrs cannot exceed 16" & vbCrLf
    If Nm("WorkDays").Value < 1 Or Nm("WorkDays").Value > 7 Then msg = msg & "- WorkDays must be 1 to 7" & vbCrLf
    For Each p In Array("AllowancePct", "IndirectPct", "AbsentPct", "OTTakeUp")
        If Nm(CStr(p)).Value < 0 Or Nm(CStr(p)).Value >= 1 Then msg = msg & "- " & p & " must be from 0% to under 100%" & vbCrLf
    Next p
    For Each p In Array("Stow", "Pick", "Pack", "Ship")
        If Nm("UPH_" & p).Value <= 0 Then msg = msg & "- UPH for " & p & " must be above zero" & vbCrLf
        If Nm("Share_" & p).Value < 0 Then msg = msg & "- Share for " & p & " cannot be negative" & vbCrLf
    Next p
    If Abs(Nm("ShareA").Value + Nm("ShareB").Value + Nm("ShareC").Value - 1) > 0.0001 Then msg = msg & "- Shift shares must total 100%" & vbCrLf
    If Application.WorksheetFunction.CountIf(Nm("ScenName"), Nm("Scenario").Value) = 0 Then msg = msg & "- Scenario is not in the scenario table" & vbCrLf
    ValidateInputs = (Len(msg) = 0)
    If Not quiet Then
        If ValidateInputs Then
            MsgBox "All inputs are valid.", vbInformation, "Input check"
        Else
            MsgBox "Fix these inputs:" & vbCrLf & vbCrLf & msg, vbExclamation, "Input check"
        End If
    End If
End Function

' Runs every scenario in the table, writes the KPIs to the comparison block on Dashboard, restores the chosen scenario.
Public Sub CompareScenarios()
    Dim orig As String, i As Long, top As Range
    If Not ValidateInputs(True) Then ValidateInputs: Exit Sub
    orig = Nm("Scenario").Value
    Set top = Nm("CompareTop")
    Application.ScreenUpdating = False
    On Error GoTo restore
    For i = 1 To Nm("ScenName").Cells.Count
        Nm("Scenario").Value = Nm("ScenName").Cells(i).Value
        Application.Calculate
        top.Offset(i - 1, 0).Value = Nm("Scenario").Value
        top.Offset(i - 1, 1).Value = Nm("KPI_Units").Value
        top.Offset(i - 1, 2).Value = Nm("KPI_Cost").Value
        top.Offset(i - 1, 3).Value = Nm("KPI_CPU").Value
        top.Offset(i - 1, 4).Value = Nm("KPI_OT").Value
        top.Offset(i - 1, 5).Value = Nm("KPI_FlexHC").Value
        top.Offset(i - 1, 6).Value = Nm("KPI_FlexDays").Value
        top.Offset(i - 1, 7).Value = Nm("KPI_Util").Value
    Next i
restore:
    Nm("Scenario").Value = orig
    Application.Calculate
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "Comparison stopped: " & Err.Description, vbExclamation
End Sub

Public Sub ExportDashboardPDF()
    Dim f As String
    f = ThisWorkbook.Path & Application.PathSeparator & "Dashboard_" & Nm("Scenario").Value & ".pdf"
    On Error GoTo fail
    ThisWorkbook.Worksheets("Dashboard").ExportAsFixedFormat Type:=xlTypePDF, Filename:=f, Quality:=xlQualityStandard
    MsgBox "Saved " & f, vbInformation
    Exit Sub
fail:
    MsgBox "Could not export: " & Err.Description, vbExclamation
End Sub
