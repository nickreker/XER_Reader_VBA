Attribute VB_Name = "Main"
Option Explicit


Declare PtrSafe Function OpenClipboard Lib "User32.dll" (ByVal hwnd As LongPtr) As LongPtr
Declare PtrSafe Function EmptyClipboard Lib "User32.dll" () As Long
Declare PtrSafe Function CloseClipboard Lib "User32.dll" () As Long
 
Public Const BREADCRUMB_TO_DASHBOARD As String = "Return to Dashboard Worksheet"
Public Const BREADCRUMB_TO_DASHBOARD_RANGE As String = "F1"
Public Const CAL_WS_NAME As String = "Calendar Report"
Public Const CAL_GEN_INFO_DEST As String = "AB16"
Public Const CALENDAR_WHOLE_AREA As String = "B3:Z39"
Public Const CALENDAR_YEAR_ADDRESS As String = "C3"
Public Const CALENDAR_STD_WORKWEEK_ADDRESS As String = "C46:X72"

Public Sub click_Do_Nothing()
End Sub
Public Sub Print_Calendar_to_Word_Click()
    Call Make_New_Word_Report
End Sub
Public Sub click_Show_User_Form()
    Call Turn_Off_Functionality
        Calendar_Picker_Form
    Call Turn_On_Functionality
End Sub

Public Sub click_New_Xer_File()
    Call Turn_Off_Functionality
        Call Make_New_Parsed_Workbook
    Call Turn_On_Functionality
End Sub

Public Sub click_Reset_Workbook()
    Call Turn_Off_Functionality
        Call Delete_All_Worksheets_but_Dash
        Call Delete_Dashboard_Tables_And_Stuff
    Call Turn_On_Functionality
End Sub

Public Sub click_New_Cal_Report()
    Call Turn_Off_Functionality
        Call Make_Cal_Rept_Worksheet
    Call Turn_On_Functionality
End Sub

Public Sub click_increment_Cal_Year()
    Call Turn_Off_Functionality
        ThisWorkbook.Worksheets(CAL_WS_NAME).Range("C3").Value = (ThisWorkbook.Worksheets(CAL_WS_NAME).Range("C3").Value + 1)
    Call Turn_On_Functionality
End Sub
Public Sub click_decrement_Cal_Year()
    Call Turn_Off_Functionality
        ThisWorkbook.Worksheets(CAL_WS_NAME).Range("C3").Value = (ThisWorkbook.Worksheets(CAL_WS_NAME).Range("C3").Value - 1)
    Call Turn_On_Functionality
End Sub

Private Sub Make_New_Parsed_Workbook()

    Call Delete_All_Worksheets_but_Dash
    Call Delete_Dashboard_Tables_And_Stuff
    
    shData.Unprotect
    
    Dim o As Split_XER
    Set o = New Split_XER
    
    o.Get_y_Read_XER
    
    Call Make_Cal_Rept_Worksheet
    Call Add_Breadcrumbs_to_Dashboard_on_all_Tabs
    
    shData.Protect
    
End Sub

Private Sub Make_Cal_Rept_Worksheet()

    Dim s As Calendar_ws
    Set s = New Calendar_ws
    
    s.Make_a_New_ws_for_a_Calendar_Report
    s.Parse_Cal_Dates
    
    Call Protect_Calendar_Report_ws
    
    shData.Activate
    shData.Range("A1").Activate
    
End Sub

Private Sub Add_Breadcrumbs_to_Dashboard_on_all_Tabs()

    Dim ws As Variant
    Dim rg As Range

    For Each ws In ThisWorkbook.Worksheets
        If ws.CodeName <> "shData" Then
            Set rg = ws.Range(BREADCRUMB_TO_DASHBOARD_RANGE)
            rg.Value2 = BREADCRUMB_TO_DASHBOARD
            
            ws.Hyperlinks.Add _
                        Anchor:=rg, _
                        Address:="", _
                        SubAddress:="'" & shData.Name & "'" & "!A1", _
                        ScreenTip:="click to Go back to the Dashboard sheet"

            rg.Font.Size = 10
            rg.Font.Italic = True
            
        End If
    Next ws
    
End Sub

'------------------------------------------------------------------------Delete/Reset Stuff

Private Sub Delete_All_Worksheets_but_Dash()

    Dim ws As Variant

    For Each ws In ThisWorkbook.Worksheets
        If ws.CodeName <> "shData" Then ws.Delete
    Next ws

End Sub

Private Sub Delete_Dashboard_Tables_And_Stuff()
    
    shData.Unprotect
    
    Dim rg As Range
    Set rg = shData.Columns("H:Z")
    
    rg.EntireColumn.Delete

    shData.Protect
    
End Sub

'------------------------------------------------------------------------Misc

Private Sub Calendar_Picker_Form()
    
    ' create the form so it initializes & fills with data
    Dim frm As Cal_Picker
    Set frm = New Cal_Picker
    
    ' figure out where to put the form
    Dim ap As Variant
    Set ap = ThisWorkbook.Application.ActiveWindow
    
    Dim fh As Integer
    fh = frm.Height
    Dim fw As Integer
    fw = frm.Width
    
    Dim x As Double
    x = ap.Left + ((ap.Width - fw) / 2)
    
    Dim y As Double
    y = ap.Top + ((ap.Height - fh) / 2)
    
    With frm
        .Left = x
        .Top = y
        .BackColor = rgbDodgerBlue
    End With
    
    ' form it up
    frm.Show vbModal
    
    ' kill kill kill
    Set frm = Nothing

End Sub

Public Sub Protect_Calendar_Report_ws()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CAL_WS_NAME)

    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:= _
        False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
        :=True, AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True
End Sub

Public Sub Unprotect_Calendar_Report_ws()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CAL_WS_NAME)
    
    ws.Unprotect
    
End Sub

'------------------------------------------------------------------------Make Word Report

Private Sub ClearClipboard()
   OpenClipboard (0&)
   EmptyClipboard
   CloseClipboard
End Sub

Private Sub Make_New_Word_Report()

    Unprotect_Calendar_Report_ws
    
    Application.CutCopyMode = True
    
    Dim wdApp As Object
    Set wdApp = CreateObject("Word.Application")

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CAL_WS_NAME)
    
    Call Delete_Crement_Buttons_to_Calendar_Report(ws)
    
    Dim rg As Range
    Set rg = ws.Range(CALENDAR_WHOLE_AREA)
    Set rg = rg.Resize(rg.Rows.Count + 6)
    
    Dim rg_exc As Range
    Set rg_exc = ws.Range(CAL_GEN_INFO_DEST).CurrentRegion
    Set rg_exc = rg_exc.Offset(rg_exc.Rows.Count + 1).CurrentRegion
    If rg_exc.Rows.Count = 1 Then GoTo Done
    Set rg_exc = rg_exc.Offset(1).Resize(rg_exc.Rows.Count - 1, 3)

    Dim d As Date
    
    d = WorksheetFunction.Min(rg_exc.Value2)
    Dim year1 As Integer
    year1 = Year(d)
    
    d = WorksheetFunction.Max(rg_exc.Value2)
    Dim year_last As Integer
    year_last = Year(d)
    
    Dim rg_wd_hrs As Range
    Set rg_wd_hrs = ws.Range(CALENDAR_STD_WORKWEEK_ADDRESS)

    With wdApp
        
        .documents.Add
        
        With .Selection
            .Font.Size = 10
            .Font.Name = "Arial"
        End With
        
        rg_wd_hrs.Copy
        .Selection.PasteSpecial DataType:=3
        .Selection.Collapse Direction:=0 'wdCollapseEnd
        .Selection.InsertBreak Type:=7 'wdPageBreak

        
        Dim y As Integer
        For y = year1 To year_last
        
            ws.Range(CALENDAR_YEAR_ADDRESS).Value2 = y
            Call ClearClipboard
            rg.Copy
            .Selection.PasteSpecial DataType:=3 'wdPasteDeviceIndependentBitmap
            .Selection.Collapse Direction:=0 'wdCollapseEnd
            If y <> year_last Then .Selection.InsertBreak Type:=7 'wdPageBreak

        Next y
        
        .Selection.Goto What:=1, Which:=1
    End With
    
    Dim year_now As Integer
    year_now = Year(Now)
    ws.Range(CALENDAR_YEAR_ADDRESS).Value2 = year_now
    
    Call Add_Crement_Buttons_to_Calendar_Report(ws)
    
    ws.Activate
    ws.Range("A1").Select
    
    wdApp.Visible = True
    
Done:
    Protect_Calendar_Report_ws
End Sub

Private Sub Delete_Crement_Buttons_to_Calendar_Report(ws As Worksheet)
    ws.Shapes("btn_Last_Year").Delete
    ws.Shapes("btn_Next_Year").Delete
    ws.Shapes("btn_Pick_Calendar").Delete
    ws.Shapes("btn_Print_Calendar").Delete
End Sub

Private Sub Add_Crement_Buttons_to_Calendar_Report(ws As Worksheet)

    Dim rg As Range
    Set rg = ws.Range("K3:L3")
    
    Dim ht As Double
    ht = rg.Height / 2
    
    Dim wd As Double
    wd = rg.Width * 0.75
    
    Dim tp As Double
    tp = rg.Top + (rg.Height / 4)
    
    ws.Buttons.Add Top:=tp, Left:=rg.Left, Height:=ht, Width:=wd
    ws.Shapes(1).Name = "btn_Last_Year"
    With ws.Shapes("btn_Last_Year")
        .DrawingObject.Caption = Chr(231) ' "ç"
        .OnAction = "click_decrement_Cal_Year"
        With .DrawingObject.Font
            .Name = "Wingdings"
            .FontStyle = "Regular"
            .Size = 14
        End With
    End With
    
    Set rg = ws.Range("P3:Q3")
    ws.Buttons.Add Top:=tp, Left:=rg.Left, Height:=ht, Width:=wd
    ws.Shapes(2).Name = "btn_Next_Year"
    With ws.Shapes("btn_Next_Year")
        .DrawingObject.Caption = Chr(232) ' "è"
        .OnAction = "click_increment_Cal_Year"
        With .DrawingObject.Font
            .Name = "Wingdings"
            .FontStyle = "Regular"
            .Size = 14
        End With
    End With
        
    Set rg = ws.Range("AB6:AC7")
    ws.Buttons.Add Top:=rg.Top, Left:=rg.Left, Height:=rg.Height, Width:=rg.Width
    ws.Shapes(3).Name = "btn_Pick_Calendar"
    With ws.Shapes("btn_Pick_Calendar")
        .DrawingObject.Caption = "Pick Calendar..."
        .OnAction = "click_Show_User_Form"
        .DrawingObject.Font.Size = 10
    End With
    
    Set rg = ws.Range("AB9:AC10")
    ws.Buttons.Add Top:=rg.Top, Left:=rg.Left, Height:=rg.Height, Width:=rg.Width
    ws.Shapes(4).Name = "btn_Print_Calendar"
    With ws.Shapes("btn_Print_Calendar")
        .DrawingObject.Caption = "Export Calendar to Word"
        .OnAction = "Print_Calendar_to_Word_Click"
        .DrawingObject.Font.Size = 10
    End With
    
End Sub

'------------------------------------------------------------------------faster faster

Private Sub Turn_Off_Functionality()
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
End Sub

Private Sub Turn_On_Functionality()
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
