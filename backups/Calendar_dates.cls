VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Calendar_dates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public raw_data As Variant
Private m_nWD As Object

Private Const LEGEND_nWD_DEST As String = "B42"
Private Const LEGEND_nWD_COLOUR As Long = 11184295 'Flatiron Light Gray
Private Const WH_GRID_COLOUR As Long = 7482624 'Flatiron Dark Blue

Private Sub Class_Initialize()
On Error GoTo error_handle

    Set m_nWD = CreateObject("Scripting.Dictionary")

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_dates" & Chr(10) & "Procedure: Class_Initialize"
End Sub

Public Property Get Non_Work_Days() As Object
On Error GoTo error_handle

    Set Non_Work_Days = m_nWD

Done:
    Exit Property

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_dates" & Chr(10) & "Property: Non_Work_Days"
End Property

Private Sub Class_Terminate()
On Error GoTo error_handle

    m_nWD.RemoveAll
    Set m_nWD = Nothing

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_dates" & Chr(10) & "Procedure: Class_Terminate"
End Sub

Public Sub Parse_Common_Info()
On Error GoTo error_handle
   
    Dim arr_temp1 As Variant
    arr_temp1 = Split(raw_data(2), "DaysOfWeek") '2 is cal_data from the idee selected
    
    Dim arr_temp2 As Variant
    arr_temp2 = Split(arr_temp1(UBound(arr_temp1)), "Exceptions")
    
    Dim arr_temp_standard_work_week As Variant
    Dim arr_temp_exceptions As Variant
    
    If (UBound(arr_temp2) - LBound(arr_temp2)) = 1 Then
        arr_temp_standard_work_week = Split(arr_temp2(LBound(arr_temp2)), "(0||")
        arr_temp_exceptions = Split(arr_temp2(UBound(arr_temp2)), "(0||")
    ElseIf (UBound(arr_temp2) = LBound(arr_temp2)) Then
        arr_temp_standard_work_week = Split(arr_temp2(LBound(arr_temp2)), "(0||")
        arr_temp_exceptions = ""
    Else
        arr_temp_standard_work_week = ""
        arr_temp_exceptions = ""
    End If
    
    If IsArray(arr_temp_standard_work_week) Then Call Parse_Standard_Work_Week(arr_temp_standard_work_week)
    
    If IsArray(arr_temp_exceptions) Then
        Dim arr As Variant
        arr = Get_Exceptions_Array(arr_temp_exceptions)
        Call Write_Exceptions_To_Cal_Report(arr)
    End If

    If UBound(raw_data) > 2 Then
        Call Get_y_Show_Global_Inherits
    End If

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_dates" & Chr(10) & "Procedure: Parse_Common_Info"
End Sub

Private Sub Get_y_Show_Global_Inherits()
On Error GoTo error_handle

    Dim arr_temp1 As Variant
    arr_temp1 = Split(raw_data(4), "Exceptions")

    Dim arr_temp_exceptions As Variant
    arr_temp_exceptions = Split(arr_temp1(UBound(arr_temp1)), "(0||")
        
    If IsArray(arr_temp_exceptions) Then
        Dim arr As Variant
        arr = Get_Exceptions_Array(arr_temp_exceptions)
    Else
        Exit Sub
    End If

    Dim dict As Object
    Set dict = Get_Dict_Exceptions_All
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CAL_WS_NAME)
    
    Dim rg As Range
    
    Dim i As Long
    For i = LBound(arr, 2) To UBound(arr, 2)
        If dict.Exists(CLng(arr(1, i))) Then
            Set rg = ws.Range(dict(CLng(arr(1, i)))).Offset(, 1)
            With rg
                .Value2 = Chr(252)
                .Font.Size = 12
                .Font.Name = "Webdings"
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
            End With
        End If
    Next i
    
    Call Add_Comment_for_Global_Inheritee

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_dates" & Chr(10) & "Procedure: Get_y_Show_Global_Inherits"
End Sub

Public Sub Add_Comment_for_Global_Inheritee()
On Error GoTo error_handle
      
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CAL_WS_NAME)
    
    Dim rg As Range
    Set rg = ws.Range(CAL_GEN_INFO_DEST).CurrentRegion
    Set rg = rg.Offset(1).Resize(rg.Rows.Count - 1)
    
    Dim cel As Variant
    For Each cel In rg.Columns(2).Cells
        If Val(raw_data(1)) = cel.Value2 Then
            cel.Offset(, 1).AddComment "Inheritor"
        ElseIf Val(raw_data(3)) = cel.Value2 Then
            cel.Style = "Note"
            cel.Offset(, 1).AddComment "Bequeather"
        End If
    Next cel

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_dates" & Chr(10) & "Procedure: Add_Comment_for_Global_Inheritee"
End Sub

Private Function Get_Dict_Exceptions_All() As Object
On Error GoTo error_handle
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CAL_WS_NAME)
    
    Dim rg As Range
    Set rg = ws.Range(CAL_GEN_INFO_DEST).CurrentRegion
    Set rg = rg.Offset(rg.Rows.Count + 1).Resize(1, 1).CurrentRegion
    Set rg = rg.Offset(1).Resize(rg.Rows.Count - 1, 1)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim cel As Variant
    For Each cel In rg.Cells
        If Not dict.Exists(cel.Value2) Then
            dict.Add key:=cel.Value2, Item:=cel.Address
        End If
    Next cel
    
    Set Get_Dict_Exceptions_All = dict

Done:
    Exit Function

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_dates" & Chr(10) & "Function: Get_Dict_Exceptions_All"
End Function

Private Sub Write_Exceptions_To_Cal_Report(arr As Variant)
On Error GoTo error_handle
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CAL_WS_NAME)
    
    Dim rg As Range
    Set rg = ws.Range(CAL_GEN_INFO_DEST).CurrentRegion
    Set rg = rg.Offset(rg.Rows.Count + 1).Resize(1, 1)
    
    rg.CurrentRegion.Clear
    
    With rg
        .Value2 = "Exceptions"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    Set rg = rg.Offset(1).Resize(UBound(arr, 2), 1)
    
    rg.Value2 = WorksheetFunction.Transpose(arr)
    rg.HorizontalAlignment = xlRight
    rg.IndentLevel = 1
    
    Dim i As Byte
    i = LBound(arr, 2)
    
    Dim j As Long
    Dim hrs() As String
    ReDim hrs(LBound(arr, 2) To UBound(arr, 2))

    For j = LBound(arr, 2) To UBound(arr, 2)
        
        If IsEmpty(arr(2, j)) Then GoTo Skip_Exception_Day
        
        Dim s As String
        s = ""
        
            For i = 2 To UBound(arr, 1)
                Dim is_even As Boolean
                is_even = (i Mod 2 = 0)
                
                If arr(i, j) = "" Then
                    GoTo Next_i_Line
                ElseIf IsEmpty(arr(i, j)) Then
                    GoTo Next_i_Line
                ElseIf i = 2 Then
                    s = CStr(arr(i, j))
                ElseIf is_even Then
                    s = s & ", " & CStr(arr(i, j))
                ElseIf Not is_even Then
                    s = s & " to " & CStr(arr(i, j))
                End If
                
                If i = UBound(arr, 1) Then GoTo Skip_Exception_Day
                
Next_i_Line:
            Next i
            
        If s = "" Then
            hrs(j) = Empty
        Else
            hrs(j) = s
        End If
        
Skip_Exception_Day:
    Next j
    
    'check if there's any hours
    Set rg = rg.Offset(, 2)
    Set rg = rg.Resize(, 1)
    
    rg.CurrentRegion.Clear
    
    Dim has_hours As Boolean
    has_hours = False
    
    Dim hr As Variant
    For Each hr In hrs
        If hr <> "" Then
            has_hours = True
            GoTo Write_Hours
        End If
    Next hr
    
    If Not has_hours Then GoTo Done
    
Write_Hours:
    If has_hours Then
        
        rg.Value2 = WorksheetFunction.Transpose(hrs)
        rg.HorizontalAlignment = xlLeft
        rg.IndentLevel = 1
        
        Set rg = rg.Offset(-1).Resize(1, 1)
        
        With rg
            .Value2 = "Hours of Exceptions"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            With .Offset(, -1).Resize(, 2).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
        
    End If

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_dates" & Chr(10) & "Procedure: Write_Exceptions_To_Cal_Report"
End Sub

Private Function Get_Exceptions_Array(e_array As Variant) As Variant
On Error GoTo error_handle

    Dim i As Long 'counter for array passed in
    
    Dim j As Byte ' counter for starts/finishes
    
    Dim k As Long ' counter for new array
    k = 0
    
    Dim arr() As Variant
    ReDim arr(1 To 49, 1 To 1)
    Dim isHours As Boolean
    isHours = False
    
    For i = LBound(e_array) To UBound(e_array)
        If InStr(e_array(i), "d|") <> 0 Then
            j = 2
            k = k + 1
            arr(1, k) = CDate(Mid(e_array(i), InStr(e_array(i), "d|") + 2, 5))
            If i <> UBound(e_array) Then
                ReDim Preserve arr(1 To 49, 1 To UBound(arr, 2) + 1)
            End If
            GoTo Continue
        End If
    
        isHours = Is_Hours(e_array(i))
        
        If isHours Then
            Dim start As Date
            start = TimeValue(Mid(e_array(i), InStr(e_array(i), "s|") + 2, 4))

            Dim fin As Date
            fin = TimeValue(Mid(e_array(i), InStr(e_array(i), "f|") + 2, 4))
            
            arr(j, k) = start
            j = j + 1
            arr(j, k) = fin
            j = j + 1
        End If
        
Continue:
    Next i
    
    Get_Exceptions_Array = arr

Done:
    Exit Function

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_dates" & Chr(10) & "Function: Get_Exceptions_Array"
End Function

Private Sub Parse_Standard_Work_Week(m_Raw_Data As Variant)
On Error GoTo error_handle

    ReDim Standard_Week(1 To 7, 1 To 5)

    Standard_Week(1, 1) = "Sunday"
    Standard_Week(2, 1) = "Monday"
    Standard_Week(3, 1) = "Tuesday"
    Standard_Week(4, 1) = "Wednesday"
    Standard_Week(5, 1) = "Thursday"
    Standard_Week(6, 1) = "Friday"
    Standard_Week(7, 1) = "Saturday"

    Dim day As Byte
    day = 1

    Dim total_hours_temp As Double
    Dim row_temp As Byte
    row_temp = 3

    Dim i As Long
    For i = LBound(m_Raw_Data) + 1 To UBound(m_Raw_Data)
    On Error Resume Next
        
        If m_Raw_Data(i) = "" Then GoTo Continue
        If IsEmpty(m_Raw_Data(i)) Then GoTo Continue
        
        Dim isWorkday As Boolean
        isWorkday = False
        If i <> UBound(m_Raw_Data) Then
            isWorkday = Prolly_A_Workday(m_Raw_Data(i)) And Is_Hours(m_Raw_Data(i + 1))
        End If

        Dim isHours As Boolean
        isHours = False
        isHours = Is_Hours(m_Raw_Data(i))

        Dim skip As Boolean
        skip = Not isHours And Not isWorkday

        If skip Then GoTo Continue
            
        If isWorkday Then
            If total_hours_temp > 0 Then Standard_Week(day, 2) = total_hours_temp
            day = (Split(m_Raw_Data(i), "(")(0))
            row_temp = 3
            total_hours_temp = 0
            GoTo Continue
        End If

        If isHours Then
        
            Dim dumbarr As Variant
            Dim dumbarr2 As Variant
            
            dumbarr = Split(m_Raw_Data(i), "s|")
            dumbarr2 = Split(dumbarr(1), "|")
            
            If (row_temp > UBound(Standard_Week, 2)) And (row_temp Mod 3 = 0) Then
                ReDim Preserve Standard_Week(1 To 7, 1 To UBound(Standard_Week, 2) + 3)
            End If

            Dim start As Date
            start = TimeValue(dumbarr2(0))
            Standard_Week(day, row_temp) = start
            row_temp = row_temp + 1

            Dim fin As Date
            fin = Split(dumbarr2(2), ")")(0)
            Standard_Week(day, row_temp) = fin
            row_temp = row_temp + 1

            Dim hrs As Double
            If start = fin Then
                hrs = 24
            ElseIf fin = #12:00:00 AM# Then
                hrs = 24 - (start * 24)
                
            Else
                hrs = (Abs((fin - start) - (start > fin))) * 24
            End If

            Standard_Week(day, row_temp) = hrs
            row_temp = row_temp + 1
            
            total_hours_temp = hrs + total_hours_temp
            
        End If
        
        If day = 7 And i = UBound(m_Raw_Data) Then Standard_Week(day, 2) = total_hours_temp

Continue:
    Next i

    Call Format_Standard_Workweek(Standard_Week)
    Call Add_Legend_Workday_Stuff(Standard_Week)

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_dates" & Chr(10) & "Procedure: Parse_Standard_Work_Week"
End Sub

Private Sub Add_Legend_Workday_Stuff(arr As Variant)
On Error GoTo error_handle

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CAL_WS_NAME)
    
    Dim rg As Range
    Set rg = ws.Range(LEGEND_nWD_DEST)
    
    With rg.Offset(, 1)
        .Value2 = "Non-Work Day"
        .HorizontalAlignment = xlLeft
        .IndentLevel = 0.5
    End With

    With rg
        .Interior.Color = LEGEND_nWD_COLOUR
        
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
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_dates" & Chr(10) & "Procedure: Add_Legend_Workday_Stuff"
End Sub

Private Sub Format_Standard_Workweek(arr As Variant)
On Error GoTo error_handle
    
    Const magic_num As Byte = 3
    Const sunday_address As String = "E48:F71"
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CAL_WS_NAME)
    
    Dim rg As Range
    Set rg = ws.Range(sunday_address)
    
    Dim i As Byte
    Dim j As Byte
    Dim col_start As Byte
    Dim row_start As Byte
    Dim col_finish As Byte
    Dim row_finish As Byte
    
    For i = LBound(arr, 1) To UBound(arr, 1)
    
        rg.Interior.Pattern = xlNone
        
        For j = magic_num To UBound(arr, 2) Step 3
            If IsEmpty(arr(i, j)) Then
                If Not m_nWD.Exists(Left(arr(i, 1), 2)) Then
                    m_nWD.Add Left(arr(i, 1), 2), Left(arr(i, 1), 2)
                End If
            Else
                
                If arr(i, j - 1) = 24 Then
                    rg.Interior.Color = WH_GRID_COLOUR
                    GoTo Next_Day
                End If

                Dim hr_start As Integer
                hr_start = CDbl(arr(i, j)) * 48
                
                Dim hr_finish As Integer
                If arr(i, j + 1) = #12:00:00 AM# Then
                    hr_finish = 48
                Else
                    hr_finish = CDbl(arr(i, j + 1)) * 48
                End If
                
                Dim k As Byte
                For k = hr_start + 1 To hr_finish
                    rg.Cells(k).Interior.Color = WH_GRID_COLOUR
                Next k

            End If
        Next j

Next_Day:

        Set rg = rg.Offset(, magic_num)
    Next i

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_dates" & Chr(10) & "Procedure: Format_Standard_Workweek"
End Sub


Private Function Prolly_A_Workday(ByVal s As String) As Boolean
On Error GoTo error_handle

    Dim p As Boolean
    p = False
    If s = "" Then GoTo Done
    If IsEmpty(s) Then GoTo Done
    p = (IsNumeric(Split(s, "(")(0))) And (InStr(s, "s|") = 0)

Done:
    Prolly_A_Workday = p
    Exit Function

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_dates" & Chr(10) & "Function: Prolly_A_Workday"
End Function

Private Function Is_Hours(ByVal s As String) As Boolean
On Error GoTo error_handle

    Dim h As Boolean
    h = False
    If s = "" Then GoTo Done
    h = (IsNumeric(Split(s, "(")(0))) And (InStr(s, "s|") <> 0) And s <> ""

Done:
    Is_Hours = h
    Exit Function

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Class Module: Calendar_dates" & Chr(10) & "Function: Is_Hours"
End Function







