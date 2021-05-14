Attribute VB_Name = "Dependencies"
Option Explicit

Public Sub Add_Conditional_Format_All()
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Dim month As String
    Dim months() As String

    months = Get_Months

    For Each ws In Worksheets
        month = Left(ws.name, 3)
        If Is_In_Array(month, months) Then
            Call Conditional_Format(ws.name)
        End If
    Next ws
    
    Application.ScreenUpdating = True
End Sub
Public Sub Conditional_Format(ByVal wsName As String)
    With Worksheets(wsName)
        .Cells.FormatConditions.Delete
        With .columns("E:E")
        
            .FormatConditions.Add Type:=xlExpression, Formula1:="=($B1=""NC"")"
            .FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 6724095
                .TintAndShade = 0
            End With
            .FormatConditions(1).StopIfTrue = False
        
            .FormatConditions.Add Type:=xlExpression, Formula1:="=OR($B1=""INPRG"", $B1=""WAPPR"")"
            .FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 6750207
                .TintAndShade = 0
            End With
            .FormatConditions(1).StopIfTrue = False
        End With

    End With
End Sub

Private Sub HideAllMonths()
    Dim ws As Worksheet
    Dim month As String
    Dim months() As String
    
    months = Active.Get_Months
    
    For Each ws In Worksheets
        month = Left(ws.name, 3)
        If (Is_In_Array(month, months)) And ws.visible = xlSheetVisible Then
            ws.visible = xlSheetVeryHidden
        End If
    Next ws
End Sub
Private Sub Format_All_Month_Tabs_With_Button()
    Dim ws As Worksheet
    Dim month As String
    Dim months() As String
    
    months = Active.Get_Months
    
    For Each ws In Worksheets
        month = Left(ws.name, 3)
        If (Is_In_Array(month, months)) And ws.visible = xlSheetVisible Then
            Call Format_Month_Tab_Button(ws.name)
        End If
    Next ws
End Sub
Public Sub Format_Month_Tab_Button(ByVal wsName As String)
    With Worksheets(wsName)
        .Rows("1:4").Interior.Color = RGB(255, 255, 255) 'Fill rows with white to remove gridlines
        .Range("A1:C1").Interior.Color = RGB(189, 215, 238) 'Refill filter header with cyan
        .Range("A5:O5").Interior.Color = RGB(189, 215, 238) 'Refill table header with cyan
        Call PositionReturnButton(wsName, "D") 'Reposition return button relative to column D
        Call AddBorder(.Range("A1:C2")) 'Add border to filter range
        Call Active.Reset_Cursor(wsName)
    End With
End Sub

' Position "Back to Dashboard" button relative to a specific column
Public Function PositionReturnButton(ByVal wsName As String, ByVal iColumns As String)
    Dim startPoint As Long
    
    With Worksheets(wsName)
        startPoint = .columns(iColumns).Left
        .Shapes("Return_Button").Left = startPoint + 18
        .Shapes("Return_Button").Top = 8
    End With
End Function
' Add all borders to specified range
Public Function AddBorder(iRange As Range)
    With iRange
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .BorderAround xlContinuous
    End With

End Function

' Return current month and year in specified format
Public Function Get_Year_Month(ByVal format As String, Optional iDate As Date) As String
    Dim returnMonth As String
    Dim returnYear As Integer
    
    ' Default to current date if no argument is supplied
    If iDate = 0 Then iDate = Date
    
    'e.g. DECEMBER 21
    If format = "mmmm-yy" Then
        returnMonth = UCase(MonthName(month(iDate)))
        returnYear = Right(year(iDate), 2)
    
        Get_Year_Month = returnMonth & " " & returnYear
    'e.g. 2021-04
    ElseIf format = "yyyy-mm" Then
        returnMonth = month(iDate)
        If Len(returnMonth) = 1 Then
            returnMonth = "0" & returnMonth
        End If
        returnYear = year(iDate)
        
        Get_Year_Month = returnYear & "-" & returnMonth
    'e.g. APR21
    ElseIf format = "mmm-yy" Then
        returnYear = Right(year(iDate), 2)
        returnMonth = UCase(MonthName(month(iDate), True))
        
        Get_Year_Month = returnMonth & returnYear
    End If

End Function

Public Function VBA_Long_To_RGB(lColor As Long) As String
    Dim iRed As Byte
    Dim iGreen As Byte
    Dim iBlue As Byte
    
    'Convert Decimal Color Code to RGB
    iRed = (lColor Mod 256)
    iGreen = (lColor \ 256) Mod 256
    iBlue = (lColor \ 65536) Mod 256
    
    'Return RGB Code
    VBA_Long_To_RGB = "RGB(" & iRed & "," & iGreen & "," & iBlue & ")"
End Function
