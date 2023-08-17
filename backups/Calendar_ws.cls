VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Calendar_ws"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CALENDAR_COLUMNS As String = "B:Z"
Private Const CALENDAR_COLUMNS_COUNT As Byte = 25
Private Const CALENDAR_START_OF_MONTHS As String = "$C$4,$K$4,$S$4,$C$13,$K$13,$S$13,$C$22,$K$22,$S$22,$C$31,$K$31,$S$31"
Private Const CALENDAR_WEEKS_PER_MONTH_COUNT As Byte = 6
Private Const CALENDAR_MONTH_NUM_ROWS As Byte = 8
Private Const CALENDAR_MONTH_NUM_COLS As Byte = 7

Private Const CALENDAR_WEEKDAY_LABELS As String = "Su,Mo,Tu,We,Th,Fr,Sa"
Private Const MONTHIES As String = "$C$6:$I$11,$K$6:$Q$11,$S$6:$Y$11,$C$15:$I$20,$K$15:$Q$20,$S$15:$Y$20,$S$24:$Y$29,$K$24:$Q$29,$C$24:$I$29,$C$33:$I$38,$K$33:$Q$38,$S$33:$Y$38"

Private Const LEGEND_nWD_DEST As String = "B42"
Private Const LEGEND_Exc_DEST As String = "H42"
Private Const LEGEND_Exc_COLOUR As Long = 11230464 'Flatiron Blue
Private Const LEGEND_Exc_Inherit_COLOUR As Long = 2797635 'Flatiron Green
Private Const LEGEND_Hr_Exc_COLOUR As Long = 2934783 'Flatiron Yellow
Private Const LEGEND_Exc_Offset_1 As Byte = 8

Private Const ROW_HEIGHT_CAL_YEAR As Byte = 35
Private Const ROW_HEIGHT_CAL_ROWS As Byte = 15

Public Sub Make_a_New_ws_for_a_Calendar_Report()
On Error GoTo error_handle

    If ThisWorkbook.Worksheets("CALENDAR").Name <> "CALENDAR" Then Exit Sub
    
    Dim ws_Cal As Worksheet
    Set ws_Cal = ThisWorkbook.Worksheets("CALENDAR")
    
    Dim sht As Variant
    For Each sht In ThisWorkbook.Worksheets
        If sht.Name = CAL_WS_NAME Then
            MsgBox "There's already a Calendar Report worksheet" & Chr(10) & "Have a nice day."
            Exit Sub
        End If
    Next sht
    
    ThisWorkbook.Worksheets.Add After:=shData

    Dim ws_New_Report As Worksheet
        
    Set ws_New_Report = ThisWorkbook.Worksheets(shData.Index + 1)
    
    With ws_New_Report
        .Activate
        .Name = CAL_WS_NAME
        .Tab.Color = rgbDodgerBlue
        Cells.Locked = False
        Cells.FormulaHidden = False
    End With
    
    ActiveWindow.DisplayGridlines = False
    
    ws_New_Report.Columns(CALENDAR_COLUMNS).ColumnWidth = 3
    
    Call Make_Std_Workweek_Grid
    Call Add_Border_Around_Whole_Calendar(ws_New_Report)
    Call Add_Formulas_and_Data_to_Calendar_Report(ws_New_Report)
    Call Add_Buttons_to_Calendar_Report(ws:=ws_New_Report)
    Call Write_Gen_Cal_Info

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Procedure: Make_a_New_ws_for_a_Calendar_Report"
End Sub


Private Sub Add_Formulas_and_Data_to_Calendar_Report(ws As Worksheet)
On Error GoTo error_handle

    Dim rg As Range
    Set rg = ws.Range(CALENDAR_YEAR_ADDRESS)
    rg.Value2 = Year(Now)
    
    Set rg = rg.Resize(1, CALENDAR_COLUMNS_COUNT - 2)
    
    With rg
        .RowHeight = ROW_HEIGHT_CAL_YEAR
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 20
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    Call Make_January(ws:=ws)
    Call Make_February_to_December(ws:=ws)

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Procedure: Add_Formulas_and_Data_to_Calendar_Report"
End Sub

Private Sub Add_Buttons_to_Calendar_Report(ws As Worksheet)
On Error GoTo error_handle

    Dim rg As Range
    Set rg = ws.Range("K3:L3")
    
    Dim ht As Double
    ht = rg.Height / 2
    
    Dim wd As Double
    wd = rg.Width * 0.75
    
    Dim tp As Double
    tp = rg.Top + (rg.Height / 4)
    
    ws.Buttons.Add Top:=tp, Left:=rg.Left, Height:=ht, Width:=wd
    ws.Shapes("Button 1").Name = "btn_Last_Year"
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
    ws.Shapes("Button 2").Name = "btn_Next_Year"
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
    ws.Shapes("Button 3").Name = "btn_Pick_Calendar"
    With ws.Shapes("btn_Pick_Calendar")
        .DrawingObject.Caption = "Pick Calendar..."
        .OnAction = "click_Show_User_Form"
        .DrawingObject.Font.Size = 10
    End With
    
    Set rg = ws.Range("AB9:AC10")
    ws.Buttons.Add Top:=rg.Top, Left:=rg.Left, Height:=rg.Height, Width:=rg.Width
    ws.Shapes("Button 4").Name = "btn_Print_Calendar"
    With ws.Shapes("btn_Print_Calendar")
        .DrawingObject.Caption = "Export Calendar to Word"
        .OnAction = "Print_Calendar_to_Word_Click"
        .DrawingObject.Font.Size = 10
    End With

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Procedure: Add_Buttons_to_Calendar_Report"
End Sub

Private Sub Make_January(ws As Worksheet)
On Error GoTo error_handle

    Dim arr_Start_of_Months As Variant
    arr_Start_of_Months = Split(CALENDAR_START_OF_MONTHS, ",")
    
    Dim rg As Range
    Set rg = ws.Range(arr_Start_of_Months(LBound(arr_Start_of_Months)))
    
    rg.FormulaR1C1 = "=DATE(R3C3," & 1 & ",1)"
    rg.NumberFormat = "mmmm"
    
    Set rg = rg.Resize(, CALENDAR_MONTH_NUM_COLS)
    
    With rg
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlBottom
        .Font.Bold = True
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        .Locked = True
        .FormulaHidden = True
    End With
    
    Set rg = rg.Offset(1)
    Dim arr() As String
    arr = Split(CALENDAR_WEEKDAY_LABELS, ",")
    
    rg.Value2 = arr
    
    With rg
        rg.HorizontalAlignment = xlCenter
        rg.VerticalAlignment = xlBottom
        rg.Font.Bold = True
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    Set rg = rg.Offset(1).Resize(CALENDAR_WEEKS_PER_MONTH_COUNT, CALENDAR_MONTH_NUM_COLS)
    Dim arr_Jan_Formulas As Variant
    arr_Jan_Formulas = Get_January_Formulas
    
    Dim i As Integer
    Dim j As Integer

    For i = LBound(arr_Jan_Formulas, 1) To UBound(arr_Jan_Formulas, 1)
        For j = LBound(arr_Jan_Formulas, 2) To UBound(arr_Jan_Formulas, 2)
             rg.Cells(i, j).FormulaR1C1 = arr_Jan_Formulas(i, j)
        Next j
    Next i

    With rg
        .NumberFormat = "d"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        .Locked = True
        .FormulaHidden = True
    End With

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Procedure: Make_January"
End Sub

Private Sub Make_February_to_December(ws As Worksheet)
On Error GoTo error_handle

    Dim arr_Start_of_Months As Variant
    arr_Start_of_Months = Split(CALENDAR_START_OF_MONTHS, ",")
    
    Dim counter_month As Byte
    
    Dim month_number As Byte
    month_number = 2
    
    For counter_month = (LBound(arr_Start_of_Months) + 1) To UBound(arr_Start_of_Months)
        Dim rg As Range
        Set rg = ws.Range(arr_Start_of_Months(counter_month))
        
        rg.FormulaR1C1 = "=DATE(R3C3," & month_number & ",1)"
        rg.NumberFormat = "mmmm"
        Set rg = rg.Resize(, CALENDAR_MONTH_NUM_COLS)
        
        With rg
            .HorizontalAlignment = xlCenterAcrossSelection
            .VerticalAlignment = xlBottom
            .Font.Bold = True
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        .Locked = True
        .FormulaHidden = True
        End With
        
        Set rg = rg.Offset(1)
        Dim arr() As String
        arr = Split(CALENDAR_WEEKDAY_LABELS, ",")
        
        rg.Value2 = arr
        rg.HorizontalAlignment = xlCenter
        rg.VerticalAlignment = xlBottom
        rg.Font.Bold = True
        
        Set rg = rg.Offset(1).Resize(CALENDAR_WEEKS_PER_MONTH_COUNT, CALENDAR_MONTH_NUM_COLS)
        
        Dim arr_Month_Formulas_non_Jan As Variant
        arr_Month_Formulas_non_Jan = Get_Month_Formulas_sans_Jan
        
        Dim i As Integer
        Dim j As Integer
    
        For i = LBound(arr_Month_Formulas_non_Jan, 1) To UBound(arr_Month_Formulas_non_Jan, 1)
            For j = LBound(arr_Month_Formulas_non_Jan, 2) To UBound(arr_Month_Formulas_non_Jan, 2)
                 rg.Cells(i, j).FormulaR1C1 = arr_Month_Formulas_non_Jan(i, j)
            Next j
        Next i
    
        With rg
            .NumberFormat = "d"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        .Locked = True
        .FormulaHidden = True
        End With
        
        month_number = month_number + 1
    Next counter_month

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Procedure: Make_February_to_December"
End Sub

Private Sub Add_Border_Around_Whole_Calendar(ws As Worksheet)
On Error GoTo error_handle

    Dim rg As Range
    Set rg = ws.Range(CALENDAR_WHOLE_AREA)
    
    With rg
        .RowHeight = ROW_HEIGHT_CAL_ROWS
        
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
    End With

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Procedure: Add_Border_Around_Whole_Calendar"
End Sub

Private Function Get_January_Formulas() As String()
On Error GoTo error_handle

    Dim arr(1 To 6, 1 To 7) As String
    
    arr(1, 1) = "=IF(WEEKDAY(DATE(R3C3,1,1),1)=MOD(1,7),DATE(R3C3,1,1),"""")"
    arr(1, 2) = "=IF(RC[-1]="""",IF(WEEKDAY(DATE(R3C3,1,1),1)=MOD(1,7)+1,DATE(R3C3,1,1),""""),RC[-1]+1)"
    arr(1, 3) = "=IF(RC[-1]="""",IF(WEEKDAY(DATE(R3C3,1,1),1)=MOD(1+1,7)+1,DATE(R3C3,1,1),""""),RC[-1]+1)"
    arr(1, 4) = "=IF(RC[-1]="""",IF(WEEKDAY(DATE(R3C3,1,1),1)=MOD(1+2,7)+1,DATE(R3C3,1,1),""""),RC[-1]+1)"
    arr(1, 5) = "=IF(RC[-1]="""",IF(WEEKDAY(DATE(R3C3,1,1),1)=MOD(1+3,7)+1,DATE(R3C3,1,1),""""),RC[-1]+1)"
    arr(1, 6) = "=IF(RC[-1]="""",IF(WEEKDAY(DATE(R3C3,1,1),1)=MOD(1+4,7)+1,DATE(R3C3,1,1),""""),RC[-1]+1)"
    arr(1, 7) = "=IF(RC[-1]="""",IF(WEEKDAY(DATE(R3C3,1,1),1)=MOD(1+5,7)+1,DATE(R3C3,1,1),""""),RC[-1]+1)"
    arr(2, 1) = "=IF(R[-1]C[6]="""","""",IF(MONTH(R[-1]C[6]+1)<>MONTH(R[-1]C[6]),"""",R[-1]C[6]+1))"
    arr(2, 2) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(2, 3) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(2, 4) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(2, 5) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(2, 6) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(2, 7) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(3, 1) = "=IF(R[-1]C[6]="""","""",IF(MONTH(R[-1]C[6]+1)<>MONTH(R[-1]C[6]),"""",R[-1]C[6]+1))"
    arr(3, 2) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(3, 3) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(3, 4) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(3, 5) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(3, 6) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(3, 7) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(4, 1) = "=IF(R[-1]C[6]="""","""",IF(MONTH(R[-1]C[6]+1)<>MONTH(R[-1]C[6]),"""",R[-1]C[6]+1))"
    arr(4, 2) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(4, 3) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(4, 4) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(4, 5) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(4, 6) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(4, 7) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(5, 1) = "=IF(R[-1]C[6]="""","""",IF(MONTH(R[-1]C[6]+1)<>MONTH(R[-1]C[6]),"""",R[-1]C[6]+1))"
    arr(5, 2) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(5, 3) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(5, 4) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(5, 5) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(5, 6) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(5, 7) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(6, 1) = "=IF(R[-1]C[6]="""","""",IF(MONTH(R[-1]C[6]+1)<>MONTH(R[-1]C[6]),"""",R[-1]C[6]+1))"
    arr(6, 2) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(6, 3) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(6, 4) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(6, 5) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(6, 6) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(6, 7) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    
    Get_January_Formulas = arr

Done:
    Exit Function

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Function: Get_January_Formulas"
End Function

Private Function Get_Month_Formulas_sans_Jan()
On Error GoTo error_handle

    Dim arr(1 To 6, 1 To 7) As String
    
    arr(1, 1) = "=IF(WEEKDAY(R[-2]C,1)=MOD(1,7),R[-2]C,"""")"
    arr(1, 2) = "=IF(RC[-1]="""",IF(WEEKDAY(R[-2]C[-1],1)=MOD(1,7)+1,R[-2]C[-1],""""),RC[-1]+1)"
    arr(1, 3) = "=IF(RC[-1]="""",IF(WEEKDAY(R[-2]C[-2],1)=MOD(1+1,7)+1,R[-2]C[-2],""""),RC[-1]+1)"
    arr(1, 4) = "=IF(RC[-1]="""",IF(WEEKDAY(R[-2]C[-3],1)=MOD(1+2,7)+1,R[-2]C[-3],""""),RC[-1]+1)"
    arr(1, 5) = "=IF(RC[-1]="""",IF(WEEKDAY(R[-2]C[-4],1)=MOD(1+3,7)+1,R[-2]C[-4],""""),RC[-1]+1)"
    arr(1, 6) = "=IF(RC[-1]="""",IF(WEEKDAY(R[-2]C[-5],1)=MOD(1+4,7)+1,R[-2]C[-5],""""),RC[-1]+1)"
    arr(1, 7) = "=IF(RC[-1]="""",IF(WEEKDAY(R[-2]C[-6],1)=MOD(1+5,7)+1,R[-2]C[-6],""""),RC[-1]+1)"
    arr(2, 1) = "=IF(R[-1]C[6]="""","""",IF(MONTH(R[-1]C[6]+1)<>MONTH(R[-1]C[6]),"""",R[-1]C[6]+1))"
    arr(2, 2) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(2, 3) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(2, 4) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(2, 5) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(2, 6) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(2, 7) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(3, 1) = "=IF(R[-1]C[6]="""","""",IF(MONTH(R[-1]C[6]+1)<>MONTH(R[-1]C[6]),"""",R[-1]C[6]+1))"
    arr(3, 2) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(3, 3) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(3, 4) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(3, 5) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(3, 6) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(3, 7) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(4, 1) = "=IF(R[-1]C[6]="""","""",IF(MONTH(R[-1]C[6]+1)<>MONTH(R[-1]C[6]),"""",R[-1]C[6]+1))"
    arr(4, 2) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(4, 3) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(4, 4) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(4, 5) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(4, 6) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(4, 7) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(5, 1) = "=IF(R[-1]C[6]="""","""",IF(MONTH(R[-1]C[6]+1)<>MONTH(R[-1]C[6]),"""",R[-1]C[6]+1))"
    arr(5, 2) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(5, 3) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(5, 4) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(5, 5) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(5, 6) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(5, 7) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(6, 1) = "=IF(R[-1]C[6]="""","""",IF(MONTH(R[-1]C[6]+1)<>MONTH(R[-1]C[6]),"""",R[-1]C[6]+1))"
    arr(6, 2) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(6, 3) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(6, 4) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(6, 5) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(6, 6) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"
    arr(6, 7) = "=IF(RC[-1]="""","""",IF(MONTH(RC[-1]+1)<>MONTH(RC[-1]),"""",RC[-1]+1))"

    Get_Month_Formulas_sans_Jan = arr

Done:
    Exit Function

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Function: Get_Month_Formulas_sans_Jan"
End Function
Private Sub Write_Gen_Cal_Info()
On Error GoTo error_handle

    Dim arr As Variant
    arr = Get_General_Cal_Info
    
    Dim chk As Boolean
    chk = Evaluate("ISREF('" & CAL_WS_NAME & "'!A1)")
    
    If chk Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(CAL_WS_NAME)
    Else
        End
    End If
    
    Dim rg As Range
    Set rg = ws.Range(CAL_GEN_INFO_DEST).CurrentRegion
    rg.ClearContents
    rg.ClearComments
    Set rg = rg.Resize(UBound(arr, 2), UBound(arr, 1))
    
    
    rg.Value2 = WorksheetFunction.Transpose(arr)
    
    Dim col As Variant
    For Each col In rg.Columns
        col.EntireColumn.AutoFit
        col.EntireColumn.HorizontalAlignment = xlCenter
    Next col
    
'    rg.HorizontalAlignment = xlCenter
    rg.Columns(3).HorizontalAlignment = xlLeft
    rg.Columns(3).IndentLevel = 1
    
    rg.Rows(1).HorizontalAlignment = xlCenter
    rg.Rows(1).Font.Bold = True
    With rg.Rows(1).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
    End With
    
    '---- MAGIC NUMBERS RULE!!--------
    '------"Fuck you future Nick" - Nick rn
    
    rg.Columns(1).EntireColumn.ColumnWidth = 13
    
    rg.Columns(2).EntireColumn.ColumnWidth = 8
        
    rg.Columns(4).EntireColumn.ColumnWidth = 8

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Procedure: Write_Gen_Cal_Info"
End Sub

'--------------------------------------------------------------------------------------------

Public Sub Parse_Cal_Dates(Optional idee As String)
On Error GoTo error_handle

    Unprotect_Calendar_Report_ws
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CAL_WS_NAME)
    
    ws.Cells.FormatConditions.Delete
    
    Dim rg As Range
    Set rg = ws.Range(CAL_GEN_INFO_DEST).CurrentRegion
    rg.ClearComments
    
    Dim s As Variant
    If idee = "" Then
        s = Get_Cal_Data_for_Date_Parsing
    Else
        s = Get_Cal_Data_for_Date_Parsing(idee)
    End If

    Dim d As Calendar_dates
    Set d = New Calendar_dates
    d.raw_data = s
    
    d.Parse_Common_Info
    
    Dim dict As Object
    Set dict = d.Non_Work_Days
    
    Call Add_Format_y_Refs_to_Cal_Rept_for_Cal_Selected(s(LBound(s)))
    Call Add_Exception_Conditional_Formats
    If Not IsEmpty(dict) Then Call Add_nWD_Conditional_Formats(dict, ws)
    
    Protect_Calendar_Report_ws

Done:
    Exit Sub

error_handle:
    Protect_Calendar_Report_ws
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Procedure: Parse_Cal_Dates"
End Sub

Private Sub Add_Exception_Conditional_Formats()
On Error GoTo error_handle

    Call Clear_Legend_of_Exception_Types

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CAL_WS_NAME)
    
    Dim rg_exc As Range
    Set rg_exc = ws.Range(CAL_GEN_INFO_DEST).CurrentRegion
    Set rg_exc = rg_exc.Offset(rg_exc.Rows.Count + 1).CurrentRegion
    If rg_exc.Rows.Count = 1 Then GoTo Done
    Set rg_exc = rg_exc.Offset(1).Resize(rg_exc.Rows.Count - 1, 3)
    
    Dim rg_m As Range
    Set rg_m = ws.Range(MONTHIES)
    
    Call Add_Exceptions_to_Legend(ws)
    
    Dim formula_string As String
    formula_string = "=AND(ISNUMBER(MATCH(C6," & rg_exc.Columns(1).Address & ",0)),C5<>"""")"
    
    With rg_m
        .FormatConditions.Add Type:=xlExpression, Formula1:=formula_string
        .FormatConditions(rg_m.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Interior.Color = LEGEND_Exc_COLOUR
        .FormatConditions(1).Font.Color = rgbWhite
        .FormatConditions(1).StopIfTrue = False
    End With

    If WorksheetFunction.CountA(rg_exc.Columns(2)) <> 0 Then
        
        Call Add_Inherits_to_Legend(ws, LEGEND_Exc_Offset_1)
        
        formula_string = "=INDEX(" & rg_exc.Columns(2).Address & ",MATCH( C6," & rg_exc.Columns(1).Address & ",0))=UNICHAR(252)"
        
        With rg_m
            .FormatConditions.Add Type:=xlExpression, Formula1:=formula_string
            .FormatConditions(rg_m.FormatConditions.Count).SetFirstPriority
            .FormatConditions(1).Interior.Color = LEGEND_Exc_Inherit_COLOUR
            .FormatConditions(1).StopIfTrue = True
        End With
        
    End If
    
    If WorksheetFunction.CountA(rg_exc.Columns(3)) <> 0 Then
    
        Dim o As Byte
        If WorksheetFunction.CountA(rg_exc.Columns(2)) <> 0 Then
            o = 0
        Else
            o = LEGEND_Exc_Offset_1
        End If
        
        Call Add_Hr_Exc_to_Legend(ws, o)

        formula_string = "=INDEX(" & rg_exc.Columns(3).Address & ",MATCH( C6," & rg_exc.Columns(1).Address & ",0))<>0"
        
        With rg_m
            .FormatConditions.Add Type:=xlExpression, Formula1:=formula_string
            .FormatConditions(rg_m.FormatConditions.Count).SetFirstPriority
            .FormatConditions(1).Interior.Color = LEGEND_Hr_Exc_COLOUR
            .FormatConditions(1).StopIfTrue = True
        End With
        
    End If

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Procedure: Add_Exception_Conditional_Formats"
End Sub

Private Sub Clear_Legend_of_Exception_Types()
On Error GoTo error_handle

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CAL_WS_NAME)
    
    Dim rg As Range
    Set rg = ws.Range(LEGEND_Exc_DEST)
    
    rg.Clear
    rg.Offset(, 1).Clear
    
    Set rg = rg.Offset(, LEGEND_Exc_Offset_1)
    
    rg.Clear
    rg.Offset(, 1).Clear
    
    Set rg = ws.Range(LEGEND_nWD_DEST).Offset(2)
    
    rg.Clear
    rg.Offset(, 1).Clear

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Procedure: Clear_Legend_of_Exception_Types"
End Sub

Private Sub Add_Exceptions_to_Legend(ws As Worksheet)
On Error GoTo error_handle

    Dim rg As Range
    Set rg = ws.Range(LEGEND_Exc_DEST)
    
    With rg.Offset(, 1)
        .Value2 = "Exception Day (Non-Work)"
        .HorizontalAlignment = xlLeft
        .IndentLevel = 0.5
    End With

    With rg
        .Interior.Color = LEGEND_Exc_COLOUR
        
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
    End With

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Procedure: Add_Exceptions_to_Legend"
End Sub
Private Sub Add_Inherits_to_Legend(ws As Worksheet, o As Byte)
On Error GoTo error_handle

    Dim rg As Range
    Set rg = ws.Range(LEGEND_Exc_DEST).Offset(, o)
    
    With rg.Offset(, 1)
        .Value2 = "Inherited Exception (i.e. From Global)"
        .HorizontalAlignment = xlLeft
        .IndentLevel = 0.5
    End With

    With rg
        .Interior.Color = LEGEND_Exc_Inherit_COLOUR
        
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
    End With

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Procedure: Add_Inherits_to_Legend"
End Sub

Private Sub Add_Hr_Exc_to_Legend(ws As Worksheet, o As Byte)
On Error GoTo error_handle

    Dim rg As Range
    
    If o <> 0 Then
        Set rg = ws.Range(LEGEND_Exc_DEST).Offset(, o)
    ElseIf o = 0 Then
        Set rg = ws.Range(LEGEND_nWD_DEST).Offset(2)
    End If
    
    With rg.Offset(, 1)
        .Value2 = "Work Hour Exception"
        .HorizontalAlignment = xlLeft
        .IndentLevel = 0.5
    End With

    With rg
        .Interior.Color = LEGEND_Hr_Exc_COLOUR
        
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
    End With

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Procedure: Add_Hr_Exc_to_Legend"
End Sub

Private Sub Add_nWD_Conditional_Formats(non_Workdays As Object, Cal_Rept_Worksheet As Worksheet)
On Error GoTo error_handle

    Const CALENDAR_WHOLE_AREA As String = "B3:Z39"
    Const CAL_WEEKDAYS_FIRST_ROW As Byte = 5
    Const CAL_WEEKDAYS_FIRST_COL As String = "C"
    
    Dim ws As Worksheet
    Set ws = Cal_Rept_Worksheet
    
    Dim rg As Range
    Set rg = ws.Range(CALENDAR_WHOLE_AREA)

    rg.Interior.Pattern = xlNone

    Set rg = ws.Range(MONTHIES)
    
    Dim key As Variant
    For Each key In non_Workdays.Keys
    
        Dim formula_string As String
        formula_string = "=" & CAL_WEEKDAYS_FIRST_COL & "$" & "5" & "=""" & key & """"
        
        With rg
            .FormatConditions.Add Type:=xlExpression, Formula1:=formula_string
            .FormatConditions(rg.FormatConditions.Count).SetFirstPriority
            .FormatConditions(1).Interior.Color = 11184295
            .FormatConditions(1).StopIfTrue = False
        End With
        
    Next key

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Procedure: Add_nWD_Conditional_Formats"
End Sub

Private Function Get_Cal_Data_for_Date_Parsing(Optional idee As String) As Variant
On Error GoTo error_handle

    Dim arr As Variant
    arr = Get_Raw_Data
    
    Dim clndr_data As Byte
    clndr_data = Get_Col_Num("clndr_data")
    
    Dim clndr_id As Byte
    clndr_id = Get_Col_Num("clndr_id")

    Dim default_flag As Byte
    default_flag = Get_Col_Num("default_flag")
    
    Dim base_clndr_id As Byte
    base_clndr_id = Get_Col_Num("base_clndr_id")

    Dim s As String
    Dim i As Integer
    
    If idee = "" Then
        idee = arr((LBound(arr, 1) + 1), clndr_id)
        GoTo IDEE_NOT_SET
    ElseIf idee <> "" Then
        GoTo IDEE_SET_ALREADY
    End If

    Dim gl As String
     
IDEE_SET_ALREADY:

    'find data for the id
    For i = (LBound(arr, 1) + 1) To UBound(arr, 1)
        If arr(i, clndr_id) = Val(idee) Then
            s = arr(i, clndr_data)
            If Not IsEmpty(arr(i, base_clndr_id)) Then gl = arr(i, base_clndr_id)
            GoTo SET_FUNC_TO_FINAL
        End If
    Next i
    
IDEE_NOT_SET:

    s = arr((LBound(arr, 1) + 1), clndr_data)
    
    'check if there's a default to switch to
    For i = (LBound(arr, 1) + 1) To UBound(arr, 1)
        If arr(i, default_flag) = "Y" Then
            s = arr(i, clndr_data)
            idee = arr(i, clndr_id)
            If Not IsEmpty(arr(i, base_clndr_id)) Then gl = arr(i, base_clndr_id)
        End If
    Next i

    GoTo SET_FUNC_TO_FINAL

SET_FUNC_TO_FINAL:

    Dim out() As String
    ReDim out(1 To 2)
    out(1) = idee
    out(2) = s
    
    If Not IsEmpty(gl) Then
        Dim gd As String
        gd = ""
        For i = (LBound(arr, 1) + 1) To UBound(arr, 1)
            If arr(i, clndr_id) = Val(gl) Then
                gd = arr(i, clndr_data)
                ReDim Preserve out(1 To 4)
                out(3) = gl
                out(4) = gd
                GoTo Done
            End If
        Next i
    End If

Done:

    Get_Cal_Data_for_Date_Parsing = out

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Procedure: Get_Cal_Data_for_Date_Parsing"
End Function

Private Sub Add_Format_y_Refs_to_Cal_Rept_for_Cal_Selected(idee As Variant)
On Error GoTo error_handle

    Const id_col As Byte = 2
    Const name_col As Byte = 3
    Const name_dest As String = "D40"
    Const name_dest_2 As String = "D72"
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CAL_WS_NAME)
    
    
    Dim rg As Range
    Set rg = ws.Range(CAL_GEN_INFO_DEST).CurrentRegion
       
    With rg.Offset(1)
        .Interior.Pattern = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Font.Color = rgbBlack
    End With
  
    Dim i As Integer
    For i = 2 To rg.Rows.Count
        If rg.Cells(i, id_col) = Val(idee) Then
               rg.Rows(i).Style = "Input"
               With ws.Range(name_dest)
                .Formula = "=" & rg.Cells(i, name_col).Address
                .Font.Bold = True
                .Font.Size = 11
                .Locked = True
                .FormulaHidden = True
               End With
               With ws.Range(name_dest_2)
                .Formula = "=" & rg.Cells(i, name_col).Address
                .Font.Bold = True
                .Font.Size = 11
                .Locked = True
                .FormulaHidden = True
               End With
            GoTo Done
        End If
    Next i

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Procedure: Add_Format_y_Refs_to_Cal_Rept_for_Cal_Selected"
End Sub

Private Function Get_Col_Num(s As String) As Byte
On Error GoTo error_handle

    Dim arr As Variant
    arr = Get_Raw_Data
    
    Dim i As Byte
    For i = LBound(arr, 2) To UBound(arr, 2)
        If arr(1, i) = s Then
            Get_Col_Num = i
            Exit Function
        End If
    Next i

Done:
    Exit Function

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Function: Get_Col_Num"
End Function

Private Function Get_Calendar_Worksheet() As Worksheet
On Error GoTo error_handle

    Dim chk As Boolean
    chk = Evaluate("ISREF('" & "CALENDAR" & "'!A1)")
    
    If chk Then
        Set Get_Calendar_Worksheet = ThisWorkbook.Worksheets("CALENDAR")
    Else
        End
    End If

Done:
    Exit Function

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Function: Get_Calendar_Worksheet"
End Function

Private Function Get_Raw_Data() As Variant
On Error GoTo error_handle

    Dim ws As Worksheet
    Set ws = Get_Calendar_Worksheet
    
    If ws Is Nothing Then End
    
    Dim rg As Range
    Set rg = Get_Calendar_Range
    
    Get_Raw_Data = rg.Value2

Done:
    Exit Function

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Function: Get_Raw_Data"
End Function

Private Function Get_Calendar_Range() As Range
On Error GoTo error_handle

    Dim ws As Worksheet
    Set ws = Get_Calendar_Worksheet

    Set Get_Calendar_Range = ws.Range("A1").CurrentRegion
    Set Get_Calendar_Range = Get_Calendar_Range.Offset(1).Resize(Get_Calendar_Range.Rows.Count - 1)

Done:
    Exit Function

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Function: Get_Calendar_Range"
End Function

Public Function Get_General_Cal_Info() As Variant
On Error GoTo error_handle

    Dim clndr_id As Byte
    clndr_id = Get_Col_Num("clndr_id")

    Dim default_flag As Byte
    default_flag = Get_Col_Num("default_flag")
    
    Dim clndr_name As Byte
    clndr_name = Get_Col_Num("clndr_name")
    
    Dim clndr_type As Byte
    clndr_type = Get_Col_Num("clndr_type")
    
    Dim info_cols() As Variant
    info_cols = Array(clndr_type, clndr_id, clndr_name, default_flag)
    
    Dim info_titles() As Variant
    info_titles = Array("Calendar Type", "ID", "Calendar Name", "Default")
    
    Dim arr As Variant
    arr = Get_Raw_Data
    
    Dim out_array() As String
    ReDim out_array(1 To 4, 1 To UBound(arr, 1))
    Dim cnt As Byte
    For cnt = LBound(out_array, 1) To UBound(out_array, 1)
        out_array(cnt, 1) = info_titles(cnt - 1)
    Next cnt

    Dim i As Integer
    Dim j As Integer
    Dim k As Byte
    
    For i = LBound(out_array, 1) To UBound(out_array, 1)
        k = info_cols(i - 1)
        For j = (LBound(out_array, 2) + 1) To UBound(out_array, 2)
            
            Select Case k
                Case clndr_type
                    Select Case arr(j, k)
                        Case "CA_Base"
                            out_array(i, j) = "Global"
                        Case "CA_Project"
                            out_array(i, j) = "Project"
                        Case "CA_Rsrc"
                            out_array(i, j) = "Resource"
                    End Select
                Case Else
                    out_array(i, j) = arr(j, k)
            End Select
        Next j
    Next i

    Get_General_Cal_Info = out_array

Done:
    Exit Function

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Function: Get_General_Cal_Info"
End Function

Private Sub Make_Std_Workweek_Grid()
On Error GoTo error_handle

    Dim rg As Range
    Set rg = ThisWorkbook.Worksheets(CAL_WS_NAME).Range(CALENDAR_STD_WORKWEEK_ADDRESS)
    
    rg.Cells(1, 1).Value = "Standard Work Week"
    With rg.Rows(1)
        .HorizontalAlignment = xlCenterAcrossSelection
        .Font.Bold = True
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
    End With
    
    Dim wds() As Variant
    wds = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
    
    Dim i As Byte
    Dim j As Byte
    j = 3
    
    Dim wd_dest As Range
    Set wd_dest = rg.Cells(2, j)
    
    For i = LBound(wds) To UBound(wds)
        With wd_dest
            .Value2 = wds(i)
            .Font.Bold = True
            .Font.Size = 9
        End With
        
        wd_dest.Resize(, 2).HorizontalAlignment = xlCenterAcrossSelection
        
        j = j + 3
        
        With wd_dest.Offset(1).Resize(24, 2)
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThick
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThick
            End With
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThick
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThick
            End With
            With .Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
        
        Set wd_dest = rg.Cells(2, j)
    Next i

    Dim rg_times As Range
    Set rg_times = rg.Cells(3, 1)
    Set rg_times = rg_times.Resize(24, 2)
    
    With rg_times
        .Font.Size = 8
        .HorizontalAlignment = xlCenterAcrossSelection
        .NumberFormat = "h:mm;@"
        .Cells(1, 1).Value2 = 0
    End With
    
    Set rg_times = rg_times.Resize(23, 1).Offset(1)
    rg_times.FormulaR1C1 = "=R[-1]C + 1/24"

    Dim nums As Variant
    nums = Array(1, 2, 5, 8, 11, 14, 17, 20)
    
    Set rg_times = rg_times.Resize(23, 22).Offset(-1)
    
    Dim rg2 As Range
    Dim num As Variant
    
    For Each num In nums
    
    Set rg2 = rg_times.Columns(num)
        
        With rg2.Borders(xlEdgeBottom)
            .LineStyle = xlDash
            .ThemeColor = 1
            .TintAndShade = -0.499984740745262
            .Weight = xlThin
        End With
        
        With rg2.Borders(xlInsideHorizontal)
            .LineStyle = xlDash
            .ThemeColor = 1
            .TintAndShade = -0.499984740745262
            .Weight = xlThin
        End With
        
    Next num
    
    Set rg = rg.Offset(-1)
    With rg.Rows(rg.Rows.Count).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_ws" & Chr(10) & "Procedure: Make_Std_Workweek_Grid"
End Sub
