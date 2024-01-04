VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComboBoxHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents cbRef As MSForms.ComboBox
Attribute cbRef.VB_VarHelpID = -1
Public txRef As MSForms.textBox
Public wsRef As Worksheet

Private Sub cbRef_Change()

    Dim selectedValue As String
    Dim tbl As ListObject
    Dim colData As Range
    Dim colResult As Range
    Dim i As Long, LastRow As Long
    
    selectedValue = cbRef.Value
    
    Set tbl = wsRef.ListObjects("TablaEquipos") ' Change to your table name
    
    Set colData = tbl.ListColumns("Equipo").DataBodyRange ' Change to your data column name
    Set colResult = tbl.ListColumns("Consumo").DataBodyRange ' Change to your result column name
    
    LastRow = wsRef.Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 2 To LastRow
        If (selectedValue) = wsRef.Cells(i, "A") Then
            txRef = wsRef.Cells(i, "B").Value
        End If
    Next i
    
End Sub
