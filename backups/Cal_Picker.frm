VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cal_Picker 
   Caption         =   "Pick a Calendar"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13080
   OleObjectBlob   =   "Cal_Picker.frx":0000
End
Attribute VB_Name = "Cal_Picker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
On Error GoTo error_handle

    Unload Me

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Form: Cal_Picker" & Chr(10) & "Procedure: btnCancel_Click"
End Sub

Private Sub btnSelect_Click()
On Error GoTo error_handle

    Dim s As String
    s = Me.listboxCals.BoundValue
    
    Dim o As Calendar_ws
    Set o = New Calendar_ws
    
    o.Parse_Cal_Dates (s)
    
    Unload Me

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Form: Cal_Picker" & Chr(10) & "Procedure: btnSelect_Click"
End Sub

Private Sub UserForm_Initialize()

    ' Fill the listbox
   Call AddDataToListbox
   
End Sub

Private Sub AddDataToListbox()
On Error GoTo error_handle

    Call Unprotect_Calendar_Report_ws
    
    ' Get the data range
    Dim rg As Range
    Set rg = GetRange
    
    ' Link the data to the ListBox
    With listboxCals
        .RowSource = rg.Address(External:=True)
        .ColumnCount = rg.Columns.Count
        .ColumnWidths = "100;0;400;80"
        .ColumnHeads = True
        .ListIndex = 0
    End With
    
    Call Protect_Calendar_Report_ws

Done:
    Exit Sub

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Form: Cal_Picker" & Chr(10) & "Procedure: AddDataToListbox"
End Sub

Private Function GetRange() As Range
On Error GoTo error_handle

    Dim chk As Boolean
    chk = Evaluate("ISREF('" & CAL_WS_NAME & "'!A1)")
    
    If chk Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(CAL_WS_NAME)
    Else
        End
    End If
    
    Dim rg As Range
    Set rg = ws.Range(CAL_GEN_INFO_DEST).CurrentRegion.Offset(1)
    If rg.Count = 1 Then End
    Set rg = rg.Resize(rg.Rows.Count - 1)
    
    If Not IsEmpty(rg) Then Set GetRange = rg

Done:
    Exit Function

error_handle:
    MsgBox Err.Description & Chr(10) & Chr(10) & "Form: Cal_Picker" & Chr(10) & "Function: GetRange"
End Function
