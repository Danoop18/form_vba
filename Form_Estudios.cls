VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub TextBox1_Change()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Dim cell As Range
    
    
    Dim Estudio As String
    Estudio = "*" & TextBox1.Text & "*"
    
    Dim tbl As ListObject
    Dim tblName As String
    
    If TextBox1.Text <> "" Then
        ' Recorre todas las tablas en la hoja de trabajo
        For Each tbl In ws.ListObjects
            tblName = tbl.Name
            
            ' Verifica si el nombre de la tabla contiene "ESTUDIO"
            If InStr(tblName, "ESTUDI") > 0 Then
                tbl.Range.AutoFilter field:=2, Criteria1:=Estudio
               End If
            
            ' Verifica si el nombre de la tabla contiene "BANCO"
            If InStr(tblName, "BANCO") > 0 Then
                tbl.Range.AutoFilter field:=1, Criteria1:=Estudio
            End If
            
            ' Verifica si el nombre de la tabla contiene "P.ESPECIAL"
            If InStr(tblName, "P.ESPECIAL") > 0 Then
                tbl.Range.AutoFilter field:=1, Criteria1:=Estudio
            End If
            ' Verifica si el nombre de la tabla contiene "P.ESPECIAL"
            If InStr(tblName, "PATOL") > 0 Then
                tbl.Range.AutoFilter field:=1, Criteria1:=Estudio
            End If
        Next tbl
    Else
        ' Limpiar todos los filtros de todas las tablas en la hoja
        For Each tbl In ws.ListObjects
            tbl.AutoFilter.ShowAllData
        Next tbl
    End If
End Sub

