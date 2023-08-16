VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Split_XER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FILE_TITLE_ON_DASHBOARD As String = "XER File Used:"
Private Const FILE_TITLE_RANGE_ON_DASHBOARD As String = "H5"

Private Const TABLES_TITLE_ON_DASHBOARD As String = "Tables Found In XER File:"
Private Const TABLES_TITLE_RANGE_ON_DASHBOARD As String = "H8"

Private Const WIDTH_OF_TABLE_COLUMN_ON_DASH As Double = 25

Private Const PROJECT_LIST_TABLE_TITLE As String = "Projects found in XER file:"
Private Const PROJECT_LIST_TABLE_TITLE_RANGE As String = "K6"

Private Const XER_ADDED_WSs_HEADERS_ROW As Byte = 2
Private Const XER_ADDED_WSs_START_DATA_ROW As Byte = 3

Public Sub Get_y_Read_XER()

Dim xer_file_path As String
    xer_file_path = Pick_Xer_File
    
    Call Add_XER_File_To_Dashboard(file_path:=xer_file_path)
        
    Dim xer_array As Variant
    xer_array = Read_Whole_XER_File(xer_file_path)
    
    Dim xer_Tables_collection As Collection
    Set xer_Tables_collection = Get_XER_Tables(xer_array)
    
    Call Dump_XER_Into_Worksheets(xer_array, xer_Tables_collection)
    
    Dim arr_project_list As Variant
    arr_project_list = Get_List_of_Projects
    
    Call Put_Projects_List_On_Dashboard(arr_project_list)
    
    shData.Activate
    shData.Range("A1").Activate

End Sub

Private Sub Put_Projects_List_On_Dashboard(list_o_projects As Variant)

    Dim rg_title As Range
    Set rg_title = shData.Range(PROJECT_LIST_TABLE_TITLE_RANGE)
    
    rg_title.Value2 = PROJECT_LIST_TABLE_TITLE
    rg_title.Font.Bold = True
    rg_title.Font.Underline = xlSingle
    
    Dim project_list_headers As Variant
    project_list_headers = Array("Project ID", "Project Name", "P6_Secret_Project_ID")
    
    Set rg_title = rg_title.Offset(2).Resize(1, (UBound(project_list_headers) + 1))

    rg_title.Value2 = project_list_headers
    rg_title.Font.Bold = True
    rg_title.HorizontalAlignment = xlCenter
    rg_title.Font.Underline = xlSingle
    
    Dim project_list_widths As Variant
    project_list_widths = Array(25, 65, 25)
    
    rg_title.ColumnWidth = project_list_widths
    
    Dim rg_projects As Range
    Set rg_projects = rg_title.Offset(1).Resize(UBound(list_o_projects, 1))

    rg_projects.Value2 = list_o_projects

End Sub

Private Function Get_List_of_Projects() As Variant

Dim rg As Range
    Set rg = ThisWorkbook.Worksheets("PROJWBS").Range("A1").CurrentRegion
    Set rg = rg.Offset(1)
    
    Const proj_node_flag As String = "proj_node_flag"
    Dim col_with_proj_flag As Long
    col_with_proj_flag = rg.Find(proj_node_flag).Column

    'wbs_short_name
    Const wbs_short_name As String = "wbs_short_name"
    Dim col_with_wbs_short_name As Long
    col_with_wbs_short_name = rg.Find(wbs_short_name).Column
    
    'wbs_name
    Const wbs_name As String = "wbs_name"
    Dim col_with_wbs_name As Long
    col_with_wbs_name = rg.Find(wbs_name).Column
    
    'proj_id
    Const proj_id As String = "proj_id"
    Dim col_with_proj_id As Long
    col_with_proj_id = rg.Find(proj_id).Column
        
    Set rg = rg.Offset(1).Resize(rg.Rows.Count - 2)
       
    Dim arr_temp As Variant
    arr_temp = rg.Value2
    
    Dim rg2 As Range
    Set rg2 = ThisWorkbook.Worksheets("PROJECT").Range("A1").CurrentRegion
    
    Dim num_of_projects As Long
    num_of_projects = rg2.Rows.Count - 2
    
    Dim arr_projects() As String
    ReDim arr_projects(1 To num_of_projects, 1 To 3)
    
    Dim i As Long
    Dim j As Long
    j = 1
    
    For i = 1 To UBound(arr_temp)
        If arr_temp(i, col_with_proj_flag) = "Y" Then
            arr_projects(j, 1) = arr_temp(i, col_with_wbs_short_name)
            arr_projects(j, 2) = arr_temp(i, col_with_wbs_name)
            arr_projects(j, 3) = arr_temp(i, col_with_proj_id)
            j = j + 1
        End If
    Next i


    Get_List_of_Projects = arr_projects

End Function

Private Function Read_Whole_XER_File(strFile As String) As Variant
    
    Open strFile For Binary Access Read As #1 'or #FreeFile
    
    Dim MyData As String
    MyData = Space$(LOF(1))
    
    Get #1, , MyData
    
    Close #1 'or blank
    
    Dim arr As Variant
    arr = Split(MyData, vbNewLine & "%T")
    
    Call QuickSort_An_Array(arr:=arr, first:=1, last:=UBound(arr))
    
    Read_Whole_XER_File = arr
    
End Function

Private Function Get_XER_Tables(xer_array As Variant) As Collection

    Dim coll_Table_names As Collection
    Set coll_Table_names = New Collection
    
    Dim i As Long
    For i = 1 To UBound(xer_array)
        coll_Table_names.Add Trim(Left(xer_array(i), InStr(xer_array(i), "%F") - 3))
    Next i
    
    Call QuickSort_A_Collection(coll_Table_names, 1, coll_Table_names.Count)
        
    Set Get_XER_Tables = coll_Table_names
        
End Function



Private Sub Add_XER_File_To_Dashboard(file_path As String)
    
    Dim rg As Range
    Set rg = shData.Range(FILE_TITLE_RANGE_ON_DASHBOARD)
    rg.Value2 = FILE_TITLE_ON_DASHBOARD
    rg.Font.Bold = True
    rg.Font.Underline = xlSingle
    
    Dim arr_temp As Variant
    arr_temp = Split(file_path, "\")
    
    Set rg = rg.Offset(1)
    rg.Value2 = arr_temp(UBound(arr_temp))
    
    shData.Hyperlinks.Add _
                        Anchor:=rg, _
                        Address:=file_path, _
                        ScreenTip:="click to Go to the source file"
    
End Sub

Private Function Pick_Xer_File() As String

    Dim fd As FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    fd.AllowMultiSelect = False
    fd.Title = "Pick the P6 XER file to Parse"
    fd.Filters.Add "P6 XER Files", "*.xer"
    
    If fd.Show = 0 Then End
    If Dir(fd.SelectedItems.Item(1)) = "" Then
        MsgBox "There was a problem with the file."
        End
    End If
    
    Pick_Xer_File = fd.SelectedItems.Item(1)

End Function

Private Sub Dump_XER_Into_Worksheets(xer_arr As Variant, xer_table_coll As Collection)
    
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim i As Long
    For i = 1 To UBound(xer_arr)
        
        Application.StatusBar = "I'm working on the (" & xer_table_coll(i) & ") table...I'm going as fast as I can."
        
        Dim sheet As Worksheet
        Set sheet = wb.Sheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        
        Dim arr_temp As Variant
        arr_temp = Split(xer_arr(i), Chr(10))

        sheet.Name = Replace(Replace(xer_table_coll(i), " ", ""), Chr(9), "")
        
        Dim rg As Range
        Set rg = sheet.Range("A1").Resize((UBound(arr_temp) + 1), 1)

'--------------------------------------------------------------------------------------
        ' Hillbilly Try/Catch Block since sometimes the CALENDAR errors
        On Error GoTo Alternate_Transpose
'            rg.Value2 = WorksheetFunction.Transpose(arr_temp) '<--- Faster way
'        GoTo Resume_Looper

Alternate_Transpose:
        Dim row As Variant
        Dim j As Long
        j = 0
        
        For Each row In rg '<--- Slower way
            row.Value = arr_temp(j)
            j = j + 1
        Next row
            
Resume_Looper:
'--------------------------------------------------------------------------------------

        rg.TextToColumns Tab:=True
        
        sheet.Columns(1).EntireColumn.Delete
        
        Call Add_Tables_to_Dashboard(total_items:=(UBound(arr_temp) - 2), _
                                                table_name:=xer_table_coll(i))
        
        Application.StatusBar = False
        Application.DisplayStatusBar = True
        
    Next i
    
    shData.Activate
    shData.Range("A1").Activate

End Sub

Private Sub Add_Tables_to_Dashboard(total_items As Long, table_name As String)
    
    Dim rg As Range
    Set rg = shData.Range(TABLES_TITLE_RANGE_ON_DASHBOARD)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(Replace(Replace(table_name, " ", ""), Chr(9), ""))
    
    Dim rg_hyper_dest As Range
    Set rg_hyper_dest = ws.Range("A1")
    
    Dim hyper_dest As String
    hyper_dest = "'" & ws.Name & "'" & "!A1"
    
    If rg.Value2 = "" Then
        rg.Value2 = TABLES_TITLE_ON_DASHBOARD
        rg.Font.Bold = True
        rg.Font.Underline = xlSingle
        rg.Columns(1).ColumnWidth = WIDTH_OF_TABLE_COLUMN_ON_DASH
        
        Set rg = shData.Range(TABLES_TITLE_RANGE_ON_DASHBOARD).Offset(1)
            rg.Value2 = Replace(Replace(table_name, " ", ""), Chr(9), "")
        
         shData.Hyperlinks.Add _
                        Anchor:=rg, _
                        Address:="", _
                        SubAddress:=hyper_dest, _
                        ScreenTip:="click to Go to the " & Replace(Replace(table_name, " ", ""), Chr(9), "") & " sheet"
                        
        Set rg = rg.Offset(, 1)
        
        rg.Value2 = total_items
        rg.NumberFormat = "#,##0"
        
    Else
        Set rg = rg.End(xlDown).Offset(1)

       rg.Value2 = Replace(Replace(table_name, " ", ""), Chr(9), "")
       
         shData.Hyperlinks.Add _
                        Anchor:=rg, _
                        Address:="", _
                        SubAddress:=hyper_dest, _
                        ScreenTip:="click to Go to the " & Replace(Replace(table_name, " ", ""), Chr(9), "") & " sheet"
                                                
        Set rg = rg.Offset(, 1)
        
        rg.Value2 = total_items
        rg.NumberFormat = "#,##0"
        
    End If

End Sub

Private Sub QuickSort_A_Collection(coll As Collection, first As Long, last As Long)
  
  Dim vCentreVal As Variant, vTemp As Variant
  
  Dim lTempLow As Long
  Dim lTempHi As Long
  lTempLow = first
  lTempHi = last
  
  vCentreVal = coll((first + last) \ 2)
  Do While lTempLow <= lTempHi
  
    Do While coll(lTempLow) < vCentreVal And lTempLow < last
      lTempLow = lTempLow + 1
    Loop
    
    Do While vCentreVal < coll(lTempHi) And lTempHi > first
      lTempHi = lTempHi - 1
    Loop
    
    If lTempLow <= lTempHi Then
    
      ' Swap values
      vTemp = coll(lTempLow)
      
      coll.Add coll(lTempHi), After:=lTempLow
      coll.Remove lTempLow
      
      coll.Add vTemp, Before:=lTempHi
      coll.Remove lTempHi + 1
      
      ' Move to next positions
      lTempLow = lTempLow + 1
      lTempHi = lTempHi - 1
      
    End If
    
  Loop
  
  If first < lTempHi Then QuickSort_A_Collection coll, first, lTempHi
  If lTempLow < last Then QuickSort_A_Collection coll, lTempLow, last
  
End Sub

Private Sub QuickSort_An_Array(arr As Variant, first As Long, last As Long)
  
  Dim vCentreVal As Variant, vTemp As Variant
  
  Dim lTempLow As Long
  Dim lTempHi As Long
  lTempLow = first
  lTempHi = last
  
  vCentreVal = arr((first + last) \ 2)
  Do While lTempLow <= lTempHi
  
    Do While arr(lTempLow) < vCentreVal And lTempLow < last
      lTempLow = lTempLow + 1
    Loop
    
    Do While vCentreVal < arr(lTempHi) And lTempHi > first
      lTempHi = lTempHi - 1
    Loop
    
    If lTempLow <= lTempHi Then
    
        ' Swap values
        vTemp = arr(lTempLow)

        arr(lTempLow) = arr(lTempHi)
        arr(lTempHi) = vTemp
      
        ' Move to next positions
        lTempLow = lTempLow + 1
        lTempHi = lTempHi - 1
      
    End If
    
  Loop
  
  If first < lTempHi Then QuickSort_An_Array arr, first, lTempHi
  If lTempLow < last Then QuickSort_An_Array arr, lTempLow, last
  
End Sub
