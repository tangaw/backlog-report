Attribute VB_Name = "Active"
Option Explicit

Public Sub Refresh_Tabs()
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Dim month As String
    Dim months() As String
    
    months = Get_Months
    
    For Each ws In ThisWorkbook.Worksheets
        month = Left(ws.name, 3)
        If (Is_In_Array(month, months)) Then
            Call Refresh_Page(ws.name)
            Call Reset_Cursor(ws.name)
        End If
    Next ws
    
    Worksheets("Dashboard").Activate
    
    Application.ScreenUpdating = True
End Sub

Public Sub Refresh_Page(ByVal wsName As String)
    Call Update_Page_Content(wsName)
    Call Sort(wsName)
    Call Filter(wsName)
End Sub

Public Sub Update_Page_Content(ByVal wsName As String)
    Call Clear_Tab_Data(wsName)
    
    With Worksheets(wsName)
        Sheets("ALL").Range("Table_Maximo_Report_Import[#All]").AdvancedFilter _
            Action:=xlFilterCopy, CriteriaRange:=.Range("A1").CurrentRegion, _
            CopyToRange:=.Range("A5:O5"), Unique:=False
    End With
    
    Application.CutCopyMode = False
End Sub
Public Sub Sort(ByVal wsName As String)
    Dim ws As Worksheet
    Set ws = Worksheets(wsName)
    
    Dim findInprg As Range
    Dim findNc As Range
    
    With ws
        Set findInprg = .Range("B:B").Find(what:="INPRG", LookIn:=xlValues, lookat:=xlWhole)
        Set findNc = .Range("B:B").Find(what:="NC", LookIn:=xlValues, lookat:=xlWhole)
    
        If Not .AutoFilterMode Then
            .Range("A5").AutoFilter
        End If
        
        With .AutoFilter.Sort
            With .SortFields
                .Clear
                If Not findInprg Is Nothing Then
                    .Add(ws.Range("E6"), xlSortOnCellColor, xlAscending, , xlSortNormal) _
                        .SortOnValue.Color = RGB(255, 255, 102)
                End If
                If Not findNc Is Nothing Then
                    .Add(ws.Range("E6"), xlSortOnCellColor, xlAscending, , xlSortNormal) _
                        .SortOnValue.Color = RGB(255, 153, 102)
                End If
                .Add2 Key:=Range("E6"), SortOn:=xlSortOnValues, Order:=xlAscending, _
                    DataOption:=xlSortNormal
            End With
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
End Sub
Public Sub Filter(ByVal wsName As String)
    Dim findInprg As Range
    Dim findWappr As Range
    Dim findNc As Range
    
    With Worksheets(wsName)
        Set findInprg = .Range("B:B").Find(what:="INPRG", LookIn:=xlValues, lookat:=xlWhole)
        Set findNc = .Range("B:B").Find(what:="NC", LookIn:=xlValues, lookat:=xlWhole)
        
        If findInprg Is Nothing And findNc Is Nothing Then
            .Tab.Color = 15518084
            Exit Sub
        Else
            .Range("B6").CurrentRegion.AutoFilter field:=2, Criteria1:=Array("INPRG", "NC") _
                , Operator:=xlFilterValues
        End If
    End With
End Sub
' Reset cursor on all month tabs to Cell "C2"
Public Sub Reset_Cursor(ByVal wsName As String)
    With Worksheets(wsName)
        If .visible = xlSheetHidden Or .visible = xlSheetVeryHidden Then
            .visible = xlSheetVisible
            .Select
            .Range("C2").Activate
            .visible = xlSheetVeryHidden
        Else
            .Select
            .Range("C2").Activate
        End If
    End With
End Sub

Public Sub Clear_All_Month_Tabs()
    Dim ws As Worksheet
    Dim months() As String
    
    months = Get_Months
    
    For Each ws In ThisWorkbook.Worksheets
        If (Is_In_Array(Left(ws.name, 3), months)) Then
            Call Clear_Tab_Data(ws.name)
        End If
    Next ws

End Sub
Public Sub Clear_Tab_Data(ByVal wsName As String)
    With Worksheets(wsName)
        On Error Resume Next
        .ShowAllData
        On Error GoTo 0
        .Range("A5").CurrentRegion.Offset(1, 0).EntireRow.Delete    'Exclude headers
    End With
    
End Sub

Public Function Get_Months() As String()
    Dim months As String
    
    months = "JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC"
    Get_Months = Split(months, ",")
End Function
Public Function Is_In_Array(stringToBeFound As String, arr As Variant) As Boolean
    Dim i As Byte
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            Is_In_Array = True
            Exit Function
        End If
    Next i
    Is_In_Array = False
End Function
