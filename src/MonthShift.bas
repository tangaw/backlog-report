Attribute VB_Name = "MonthShift"
Option Explicit

Public Sub Add_Backlog_Month(ByVal iMonth As String, ByVal iYear As Integer)
    Dim previousMonth As Date   ' Adding previous month due to backlog inclusion
    Dim newMonth As Date   ' Use this month as template for new tab
    Dim refMonthYearTabName As String
    Dim newMonthYearTabName As String
    Dim newMonthYearPivot As String
    
    Dim newSheet As Worksheet
    
    previousMonth = DateAdd("m", -1, DateValue("01 " & iMonth & " " & iYear))
    newMonth = DateValue("01 " & iMonth & " " & iYear)
    refMonthYearTabName = Dependencies.Get_Year_Month("mmm-yy", previousMonth)
    newMonthYearTabName = Dependencies.Get_Year_Month("mmm-yy", newMonth)
    newMonthYearPivot = Dependencies.Get_Year_Month("yyyy-mm", newMonth)
    
    ' Exit sub if sheet already exists
    On Error Resume Next
        Set newSheet = Worksheets(newMonthYearTabName)
    On Error GoTo 0
    If Not newSheet Is Nothing Then
        MsgBox "Sheet already exists!"
        Exit Sub
    End If
    
    ' Clear reference sheet and copy as new sheet
    Call Active.Clear_Tab_Data(refMonthYearTabName)
    Worksheets(refMonthYearTabName).Copy after:=Worksheets(refMonthYearTabName)
    
    With Worksheets(refMonthYearTabName & " (2)")
        .name = newMonthYearTabName
        .Range("C2").Value = newMonthYearPivot
    End With
    
    Call Update_All_Pivots(newMonthYearPivot)
End Sub

' Update MonthPivot to include new month
Public Sub Update_All_Pivots(ByVal iMonthYear As String)
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim findMonthYear As Range
    Dim pivotName As String
    
    For Each pivotName In Array("MonthPivot", "SitePivot", "PiePivot")
        Set pt = Worksheets(pivotName).PivotTables(pivotName)
        Set pf = pt.PivotFields("year-month")
        pf.PivotItems(iMonthYear).visible = True
    Next pivot
    
    With Worksheets("MonthPivot")
        ' Update MonthPivot chart reference table to include new month
        Set findMonthYear = .columns("G").Find(what:=iMonthYear, LookIn:=xlValues, lookat:=xlWhole)
        If Not findMonthYear Is Nothing Then
            MsgBox "Year-month already included in chart reference table!"
            Exit Sub
        Else
            .Range("G4").End(xlDown).Offset(1, 0).Value = iMonthYear
        End If
    End With
End Sub

