Attribute VB_Name = "Dashboard"
Option Explicit

Public Sub cbLoadMonths()
    Dim ws As Worksheet
    Dim month As String
    Dim months() As String
    Dim currentMonthYear As String
    
    With Worksheets("Dashboard")
        With .cbViewMonth
            .Clear
            months = Get_Months
            currentMonthYear = Dependencies.Get_Year_Month("mmm-yy")
            
            For Each ws In Worksheets
                month = Left(ws.name, 3)
                If (Is_In_Array(month, months)) And ws.name <> currentMonthYear Then  'Exclude current year-month
                    .AddItem ws.name
                End If
            Next ws
        End With
    End With
End Sub
' Load set of predetermined site names into the view site combo box
Public Sub cbLoadSites()
    Dim sites() As String
    Dim site As Variant
    
    sites = Split("ARQUES,BOWERS,SCOTT,SITEOPS,CSR,FLSS", ",")
    
    With Worksheets("Dashboard").cbViewSite
        For Each site In sites
            .AddItem site
        Next site
    End With
End Sub

Public Sub Disable_View_Buttons()
    Dim button As Variant
    Dim arrow As String
    
    With Worksheets("Dashboard")
        For Each button In Array("ViewMonthButton", "ViewSiteButton")
            arrow = button + "Arrow"
            .Shapes(button).OnAction = ""
            .Shapes(arrow).Fill.Transparency = 0.7
        Next button
    End With
End Sub
' Sub that activates when the "View tab" buttons are pressed
Public Sub View_Button(ByVal category As String)
    Dim cbSelection As String
    Dim selectSheet As Worksheet
    
    Select Case category
        Case "Month"
            cbSelection = Worksheets("Dashboard").cbViewMonth.Value
            
            ' Catch error when entered sheet does not exist
            On Error Resume Next
            Set selectSheet = Worksheets(cbSelection)
            On Error GoTo 0
            
            If selectSheet Is Nothing Then
                MsgBox "Please select valid sheet name."
                Exit Sub
            Else
                With Worksheets(cbSelection)
                    .visible = xlSheetVisible
                    .Activate
                End With
            End If
        Case "Site"
            cbSelection = Worksheets("Dashboard").cbViewSite.Value
            
            ' Catch error
            If Not Is_In_Array(cbSelection, _
                Array("ARQUES", "BOWERS", "SCOTT", "SITEOPS", "CSR", "FLSS")) Then
                    MsgBox "Please select valid sheet name."
                    Exit Sub
            Else
                Dim monthYears() As String
                Dim allMonths As String
                Dim iCell As Range
                ' xlFilterValues only accepts String Arrays
                With Worksheets("MonthPivot")
                    For Each iCell In .Range(.Range("G5"), .Range("G5").End(xlDown))
                        allMonths = allMonths + iCell.Value + ","
                    Next iCell
                    monthYears = Split(Left(allMonths, Len(allMonths) - 1), ",")
                End With
    
                With Sheet29
                    .name = cbSelection
                    .Range("A2").Value = cbSelection
                    Call Active.Refresh_Page(.name)
                    .Range("B6").CurrentRegion.AutoFilter field:=11, Criteria1:=monthYears, Operator:=xlFilterValues
                    .visible = xlSheetVisible
                    .Activate
                    .Range("A2").Select
                End With
            End If
    End Select

End Sub

Public Sub ReturnToDashboard(Optional ByVal iType As String)
    If iType = "SiteSelect" Then
        ActiveSheet.name = "SiteSelect"
    End If
    ActiveSheet.visible = xlSheetVeryHidden
    Worksheets("Dashboard").Select
End Sub

' Toggle Pivot Design Tabs visibility
Private Sub Toggle_Design_Tabs_Visibility()
    Dim pivots() As Variant
    Dim i As Byte
    
    pivots = Array("MonthPivot", "SitePivot", "PiePivot", "ALL")
    For i = LBound(pivots) To UBound(pivots)
        With Worksheets(pivots(i))
            If .visible = xlSheetVeryHidden Then
                .visible = xlSheetVisible
            Else
                .visible = xlSheetVeryHidden
            End If
        End With
    Next i
End Sub

Public Sub ToggleSiteSlicerInfoBox()
    Dim infoButton As Shape
    Dim infoBox As Shape
    
    Set infoButton = ActiveSheet.Shapes("Info_Button_Site_Slicer")
    Set infoBox = ActiveSheet.Shapes("Info_Box_Site_Slicer")
    
    If infoButton.Fill.ForeColor.RGB = RGB(255, 243, 185) Then
        Call ToggleInfoBoxVisibility(infoButton, infoBox, False)
    Else
        Call ToggleInfoBoxVisibility(infoButton, infoBox, True)
    End If

End Sub
Public Function ToggleInfoBoxVisibility(button As Object, infoBox As Object, visible As Boolean)
    If visible = True Then
        With ActiveSheet
            button.Fill.ForeColor.RGB = RGB(255, 243, 185)
            infoBox.visible = True
        End With
    Else
        With ActiveSheet
            button.Fill.ForeColor.RGB = RGB(255, 255, 255)
            infoBox.visible = False
        End With
    End If
End Function
