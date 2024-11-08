Sub LIMPIAR()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Limpiar valores en la tabla "Estudios"
    ws.ListObjects("Estudios").ListColumns("ESTUDIOS REALIZADOS").DataBodyRange.FormulaR1C1 = "0"
    
    ' Limpiar valores en la tabla "P.ESPECIALES"
    'ws.ListObjects("P.ESPECIALES").ListColumns(2).DataBodyRange.FormulaR1C1 = "0"
     
    ws.Range("F4").Value = "HOSPITAL" ' Valor inicial
    
    ' Mostrar mensaje y capturar la respuesta del usuario
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("Listas en 0, deseas registrar", vbYesNo + vbDefaultButton1)
    
    If respuesta = vbYes Then
        UserForm1.Show ' Mostrar el formulario de registro
    Else
    ' Desplegar lista de validación en la celda F4
     With ws.Range("F5").Validation
        .Delete ' Eliminar cualquier validación existente
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="HOSPITAL, CLINICA, LABORATORIO" ' Lista de opciones
    End With
        Exit Sub ' Salir del procedimiento
    End If
End Sub


Sub CrearLibroNuevo()
    Dim newWorkbook As Workbook
    Dim templateSheet As Worksheet
    Dim newSheet As Worksheet

    ' Asignar la hoja de plantilla a una variable
    Set templateSheet = ThisWorkbook.Sheets("Template")

    ' Crear un nuevo libro de trabajo
    Set newWorkbook = Workbooks.Add

    ' Crear la hoja "GENERAL" y copiar contenido del template
    Set newSheet = newWorkbook.Sheets.Add(After:=newWorkbook.Sheets(newWorkbook.Sheets.count))
    newSheet.Name = "GENERAL"
    templateSheet.Range("A:L").Copy Destination:=newSheet.Range("A1")

    ' Crear la hoja "ASEGURADORAS" y copiar contenido del template
    Set newSheet = newWorkbook.Sheets.Add(After:=newWorkbook.Sheets(newWorkbook.Sheets.count))
    newSheet.Name = "ASEGURADORAS"
    templateSheet.Range("A:L").Copy Destination:=newSheet.Range("A1")

    ' Guardar el nuevo libro en formato XLSX
    newWorkbook.SaveAs Filename:="LISTA DE ESTUDIOS REALIZADOS CONCENTRADO costos DESGLOSE " & templateSheet.Range("D4") & " SERVICIOS 2024.xlsx", FileFormat:=xlOpenXMLWorkbook

    ' Cerrar el nuevo libro
    newWorkbook.Close SaveChanges:=True
End Sub

Sub AgregarHoja()
    Dim newSheet As Worksheet
    Dim templateSheet As Worksheet
    Dim estudiosT As Worksheet
    Dim targetSheet As Worksheet
    Dim sheetName As String
    Dim newWorkbook As Workbook
    
    ' Asignar las hojas a variables para mejorar la legibilidad
    Set templateSheet = ThisWorkbook.Sheets("Template")
    Set UVEHSheet = ThisWorkbook.Sheets("UVEH")
    Set estudiosT = ThisWorkbook.Sheets("EstudiosT")

    ' Abrir el libro creado anteriormente
    Set newWorkbook = Workbooks.Open("LISTA DE ESTUDIOS REALIZADOS CONCENTRADO costos DESGLOSE " & templateSheet.Range("D4") & " SERVICIOS 2024.xlsx")

    ' Obtener el valor de la celda F5
    sheetName = estudiosT.Range("F5").Value
If sheetName = "" Then
    MsgBox "Por favor, elija una nueva subcategoría."
    
Else
      
    ' Determinar la hoja después de la cual agregar la nueva hoja
    Select Case sheetName
        Case "HOSPITAL", "EXTERNO", "URGENCIAS"
            Set targetSheet = newWorkbook.Sheets("GENERAL")
        Case "IMSS", "INVESCLINIC", "LAPI", "ORORA", "S. MONTERREY", "SEG_ATLAS", "UNIVERSIDAD_GTO", "MAC_CELAYA", "PROTEC", "MEDICA_M", "RED_MEDICA", "M._MOVIL", "P&G", "PEMEX", "UVEH", "MED_INSURGENTES", "METLIFE", "AXXA", "S_OCUP", "STO._DOMINGO", "RYOBI", "MMS SANTANDER", "PRESTIGE TADASHI"
            Set targetSheet = newWorkbook.Sheets("ASEGURADORAS")
        Case "SUBROGADOS", "LANS", "STO._DOMINGO"
            Set targetSheet = newWorkbook.Sheets("REPORTE MENSUAL")
        Case Else
            MsgBox "Valor no válido en la celda F5"
            Exit Sub
    End Select

    ' Agregar una nueva hoja después de la hoja target
    Set newSheet = newWorkbook.Sheets.Add(After:=targetSheet)
    newSheet.Name = sheetName
    
    If sheetName = "UVEH" Then
    UVEHSheet.Range("A:L").Copy Destination:=newSheet.Range("A1")
    Else
    ' Copiar el rango "A:L" de la hoja "Template" a la nueva hoja
    templateSheet.Range("A:L").Copy Destination:=newSheet.Range("A1")

    ' Copiar y pegar valores de "Estudios[ESTUDIOS REALIZADOS]" a partir de "C7" en la nueva hoja
    Dim lastRow As Long
    lastRow = estudiosT.Range("Estudios[ESTUDIOS REALIZADOS]").Rows.count
    newSheet.Range("C7").Resize(lastRow).Value = estudiosT.Range("Estudios[ESTUDIOS REALIZADOS]").Value

    ' Guardar y cerrar el libro
    newWorkbook.Close SaveChanges:=True
    End If
    ThisWorkbook.Activate
    estudiosT.Activate
    estudiosT.Range("F5").Select ' Asegúrate de seleccionar la celda para que el usuario vea la lista de validación
    estudiosT.Range("F5").Validation.InCellDropdown = True ' Asegúrate de que la lista se despliegue
 End If
 Exit Sub
 End Sub

Dim selectedTable As String
Dim selectedTable2 As String

Function ListObjectExists(ws As Worksheet, tableName As String) As Boolean
    On Error Resume Next
    ListObjectExists = Not ws.ListObjects(tableName) Is Nothing
    On Error GoTo 0
End Function

Public dict As Object

Public Sub Initialize()
    Set dict = CreateObject("Scripting.Dictionary")
End Sub

Sub FILTRADO()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim filtroAplicado As Boolean
    
    ' Establece la hoja en la que deseas buscar las tablas
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Itera a través de todas las tablas en la hoja
    For Each tbl In ws.ListObjects
        filtroAplicado = False
        
        ' Verifica si la tabla tiene filtros aplicados
        On Error Resume Next
        filtroAplicado = tbl.AutoFilter.FilterMode
        On Error GoTo 0
        
        If filtroAplicado Then
            ' Limpia los filtros de la tabla
            tbl.AutoFilter.ShowAllData
        Else
            ' Aplica los filtros según los criterios específicos
            If InStr(1, tbl.Name, "ESTUDIO") > 0 Then
                ' Filtra la tabla en función del criterio en el Campo 1 (primera columna)
                tbl.Range.AutoFilter field:=2, Criteria1:=">0", Operator:=xlAnd
            ElseIf InStr(1, tbl.Name, "BANCO") > 0 Then
                ' Filtra la tabla en función del criterio en el Campo 2 (segunda columna)
                tbl.Range.AutoFilter field:=2, Criteria1:=">0", Operator:=xlAnd
            ElseIf InStr(1, tbl.Name, "P.ESPECIAL") > 0 Then
                ' Filtra la tabla en función del criterio en el Campo 2 (segunda columna)
                tbl.Range.AutoFilter field:=2, Criteria1:=">0", Operator:=xlAnd
            ElseIf InStr(1, tbl.Name, "PATOL") > 0 Then
                ' Filtra la tabla en función del criterio en el Campo 2 (segunda columna)
                tbl.Range.AutoFilter field:=2, Criteria1:=">0", Operator:=xlAnd
            End If
        End If
    Next tbl
End Sub

Sub REGISTRO()
UserForm1.Show
End Sub
Sub CleanUpWorksheet()
    Dim ws As Worksheet
    Dim firstLabRow As Long
    Dim lastUrgenciaRow As Long
    Dim labFound As Boolean
    Dim urgenciaFound As Boolean
    
    Dim i As Long, j As Long
    Dim cell As Range

    Set ws = ThisWorkbook.ActiveSheet ' Ajusta esto si necesitas apuntar a una hoja específica

    ' Inicializa las variables
    firstLabRow = 0
    lastUrgenciaRow = 0
    labFound = False
    urgenciaFound = False

    ' Busca las ocurrencias de "laboratorio" y "servicio de urgencias" en todas las celdas de la hoja
    For i = 1 To ws.Cells(ws.Rows.count, 1).End(xlUp).Row
        For j = 1 To ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
            Set cell = ws.Cells(i, j)
            If InStr(1, LCase(cell.Value), "laboratorio") > 0 Then
                If Not labFound Then
                    firstLabRow = i
                    labFound = True
                End If
            End If
            
            If InStr(1, LCase(cell.Value), "servicio de urgencias") > 0 Then
                lastUrgenciaRow = i
                urgenciaFound = True
            End If
        Next j
    Next i

    ' Si se encuentran ambas palabras
    If labFound And urgenciaFound Then
        ' Eliminar filas por encima de la primera ocurrencia de "laboratorio"
        If firstLabRow > 1 Then
            ws.Rows("1:" & firstLabRow - 1).Delete
        End If

        ' Ajustar lastUrgenciaRow después de la eliminación inicial
        lastUrgenciaRow = lastUrgenciaRow - (firstLabRow - 1)
        
        ' Eliminar filas por debajo de la última ocurrencia de "servicio de urgencias"
        If lastUrgenciaRow < ws.Rows.count Then
            ws.Rows(lastUrgenciaRow + 1 & ":" & ws.Rows.count).Delete
        End If
    Else
        MsgBox "No se encontraron ambas palabras en la hoja.", vbExclamation
    End If
End Sub



Sub GENERAL2()
    ' Variables para almacenar las sumas
    Dim estRows As Long
    Dim estRows2 As Long
    Dim estRows3 As Long
    Dim estRows4 As Long
    ' Sumar valores de celdas específicas en diferentes hojas y colocar la fórmula en la celda activa
      ActiveSheet.Range("C8").Formula = "=SUM(HOSPITAL!C8,EXTERNO!C8,URGENCIAS!C8,ASEGURADORAS!C8)"
    ' Rellenar el rango "ESTUDIO57[Columna1]" con la fórmula
    estRows = ActiveSheet.ListObjects(4).ListRows.count
    ActiveSheet.Range("C8").AutoFill Destination:=ActiveSheet.Range("C8:C" & 8 + estRows - 1)
    
    ' Sumar valores de celdas específicas de la hoja "BANCO"
       ActiveSheet.Range("D130").Formula = "=SUM(HOSPITAL!D130,EXTERNO!D130,URGENCIAS!D130,ASEGURADORAS!D130)"
       estRows2 = ActiveSheet.ListObjects(3).ListRows.count
    ActiveSheet.Range("D130").AutoFill Destination:=ActiveSheet.Range("D130:D" & 130 + estRows2 - 1)
               
    ' Colocar la suma en la celda D130 de la hoja activa
   ActiveSheet.Range("D143").Formula = "=SUM(HOSPITAL!D143,EXTERNO!D143,URGENCIAS!D143,ASEGURADORAS!D143)"
   estRows3 = ActiveSheet.ListObjects(2).ListRows.count
    ActiveSheet.Range("D143").AutoFill Destination:=ActiveSheet.Range("D143:D" & 143 + estRows3 - 1)
                   
    ActiveSheet.Range("D288").Formula = "=SUM(HOSPITAL!D288,EXTERNO!D288,URGENCIAS!D288,ASEGURADORAS!D288)"
    estRows4 = ActiveSheet.ListObjects(1).ListRows.count
    ActiveSheet.Range("D288").AutoFill Destination:=ActiveSheet.Range("D288:D" & 288 + estRows4 - 1)
               
    ' Mostrar mensaje de finalización
    MsgBox "Resumen completado, revisa que los estudios especiales coincidan"
End Sub
Sub ASEGURADORAS()
    Dim sheetsArray As Variant
    Dim sheetName As String
    Dim sumFormula As String
    Dim estRows As Long
    Dim i As Long
'SUMA(IMSS!C8,PEMEX!C8,S. MONTERREY!C8,'P&G'!C8,LAPI!C8,INVESCLINIC!C8,ORORA!C8)
    ' Array de nombres de hojas
    sheetsArray = Array("IMSS", "PEMEX", "'S. MONTERREY'", " 'P&G'", "LAPI", "INVESCLINIC", "ORORA")
    ' Construir la fórmula SUM para la celda C8
    sumFormula = "=SUM("
    For i = LBound(sheetsArray) To UBound(sheetsArray)
        sheetName = sheetsArray(i)
        On Error Resume Next
        If Not IsError(Worksheets(sheetName)) Then
            sumFormula = sumFormula & sheetName & "!C8,"
        End If
        On Error GoTo 0
    Next i
    ' Eliminar la última coma y cerrar el paréntesis
    sumFormula = Left(sumFormula, Len(sumFormula) - 1) & ")"
    
    ' Insertar la fórmula en la celda C8 de la hoja activa
    ActiveSheet.Range("C8").Formula = sumFormula
    
    ' Determinar el número de filas en la tabla ESTUDIO57
    estRows = ActiveSheet.ListObjects(4).ListRows.count
    
    ' Rellenar el rango desde C8 hacia abajo hasta el número de filas en la tabla ESTUDIO57
    ActiveSheet.Range("C8").AutoFill Destination:=ActiveSheet.Range("C8:C" & 8 + estRows - 1), Type:=xlFillValues
    
    ' Repetir el proceso para las demás tablas y celdas específicas
    ' Construir la fórmula SUM para la celda D130
    sumFormula = "=SUM("
    For i = LBound(sheetsArray) To UBound(sheetsArray)
        sheetName = sheetsArray(i)
        On Error Resume Next
        If Not IsError(Worksheets(sheetName)) Then
            sumFormula = sumFormula & sheetName & "!D130,"
        End If
        On Error GoTo 0
    Next i
    ' Eliminar la última coma y cerrar el paréntesis
    sumFormula = Left(sumFormula, Len(sumFormula) - 1) & ")"
    
    ' Insertar la fórmula en la celda D130 de la hoja activa
    ActiveSheet.Range("D130").Formula = sumFormula
    estRows = ActiveSheet.ListObjects(3).ListRows.count
    ActiveSheet.Range("D130").AutoFill Destination:=ActiveSheet.Range("D130:D" & 130 + estRows - 1), Type:=xlFillValues
    
    ' Construir la fórmula SUM para la celda D143
    sumFormula = "=SUM("
    For i = LBound(sheetsArray) To UBound(sheetsArray)
        sheetName = sheetsArray(i)
        On Error Resume Next
        If Not IsError(Worksheets(sheetName)) Then
            sumFormula = sumFormula & sheetName & "!D143,"
        End If
        On Error GoTo 0
    Next i
    sumFormula = Left(sumFormula, Len(sumFormula) - 1) & ")"
    
    ' Insertar la fórmula en la celda D143 de la hoja activa
    ActiveSheet.Range("D143").Formula = sumFormula
    estRows = ActiveSheet.ListObjects(2).ListRows.count
    ActiveSheet.Range("D143").AutoFill Destination:=ActiveSheet.Range("D143:D" & 143 + estRows - 1), Type:=xlFillValues
    
    ' Construir la fórmula SUM para la celda D288
    sumFormula = "=SUM("
    For i = LBound(sheetsArray) To UBound(sheetsArray)
        sheetName = sheetsArray(i)
        On Error Resume Next
        If Not IsError(Worksheets(sheetName)) Then
            sumFormula = sumFormula & sheetName & "!D288,"
        End If
        On Error GoTo 0
    Next i
    sumFormula = Left(sumFormula, Len(sumFormula) - 1) & ")"
    
    ' Insertar la fórmula en la celda D288 de la hoja activa
    ActiveSheet.Range("D288").Formula = sumFormula
    estRows = ActiveSheet.ListObjects(1).ListRows.count
    ActiveSheet.Range("D288").AutoFill Destination:=ActiveSheet.Range("D288:D" & 288 + estRows - 1), Type:=xlFillValues
    
    ' Mostrar mensaje de finalización
    MsgBox "Resumen completado, revisa que los estudios especiales coincidan"
End Sub
Sub GENERAL()
    Dim sheetsArray As Variant
    Dim sheetName As String
    Dim sumFormula As String
    Dim estRows As Long
    Dim i As Long

    sheetsArray = Array("HOSPITAL", "EXTERNO", "URGENCIAS", "ASEGURADORAS")

' Construir la fórmula SUM para la celda C8
sumFormula = "=SUM("
For i = LBound(sheetsArray) To UBound(sheetsArray)
    sheetName = sheetsArray(i)
    On Error Resume Next
   If ThisWorkbook.Sheets(sheetName) Is Nothing Then
    ' La hoja no existe en el libro actual, por lo que no se hace nada y se pasa a la siguiente iteración
Else
         sumFormula = sumFormula & "'" & sheetName & "'!C8,"
    End If
    On Error GoTo 0
Next i
' Eliminar la última coma y cerrar el paréntesis
sumFormula = Left(sumFormula, Len(sumFormula) - 1) & ")"

' Insertar la fórmula en la celda C8 de la hoja activa
ActiveSheet.Range("C7").Formula = sumFormula

' Determinar el número de filas en la tabla ESTUDIO57 si existe
If Not IsError(ActiveSheet.ListObjects(4)) Then
    estRows = ActiveSheet.ListObjects(4).ListRows.count

    ' Rellenar el rango desde C8 hacia abajo hasta el número de filas en la tabla ESTUDIO57
    ActiveSheet.Range("C8").AutoFill Destination:=ActiveSheet.Range("C8:C" & 8 + estRows - 1), Type:=xlFillValues
End If
  
    ' Repetir el proceso para las demás tablas y celdas específicas
    ' Construir la fórmula SUM para la celda D130
    sumFormula = "=SUM("
    For i = LBound(sheetsArray) To UBound(sheetsArray)
        sheetName = sheetsArray(i)
        On Error Resume Next
        If ThisWorkbook.Sheets(sheetName) Is Nothing Then
    ' La hoja no existe en el libro actual, por lo que no se hace nada y se pasa a la siguiente iteración
Else
            sumFormula = sumFormula & "'" & sheetName & "'!D130,"
        End If
        On Error GoTo 0
    Next i
    ' Eliminar la última coma y cerrar el paréntesis
    sumFormula = Left(sumFormula, Len(sumFormula) - 1) & ")"
    
    ' Insertar la fórmula en la celda D130 de la hoja activa
    ActiveSheet.Range("D130").Formula = sumFormula
    estRows = ActiveSheet.ListObjects(3).ListRows.count
    ActiveSheet.Range("D130").AutoFill Destination:=ActiveSheet.Range("D130:D" & 130 + estRows - 1), Type:=xlFillValues
    
    ' Construir la fórmula SUM para la celda D143
    sumFormula = "=SUM("
    For i = LBound(sheetsArray) To UBound(sheetsArray)
        sheetName = sheetsArray(i)
        On Error Resume Next
        If ThisWorkbook.Sheets(sheetName) Is Nothing Then
    ' La hoja no existe en el libro actual, por lo que no se hace nada y se pasa a la siguiente iteración
Else
            sumFormula = sumFormula & "'" & sheetName & "'!D143,"
        End If
        On Error GoTo 0
    Next i
    sumFormula = Left(sumFormula, Len(sumFormula) - 1) & ")"
    
    ' Insertar la fórmula en la celda D143 de la hoja activa
    ActiveSheet.Range("D143").Formula = sumFormula
    estRows = ActiveSheet.ListObjects(2).ListRows.count
    ActiveSheet.Range("D143").AutoFill Destination:=ActiveSheet.Range("D143:D" & 143 + estRows - 1), Type:=xlFillValues
    
    ' Construir la fórmula SUM para la celda D288
    sumFormula = "=SUM("
    For i = LBound(sheetsArray) To UBound(sheetsArray)
        sheetName = sheetsArray(i)
        On Error Resume Next
        If ThisWorkbook.Sheets(sheetName) Is Nothing Then
    ' La hoja no existe en el libro actual, por lo que no se hace nada y se pasa a la siguiente iteración
Else
            sumFormula = sumFormula & "'" & sheetName & "'!D288,"
        End If
        On Error GoTo 0
    Next i
    sumFormula = Left(sumFormula, Len(sumFormula) - 1) & ")"
    
    ' Insertar la fórmula en la celda D288 de la hoja activa
    ActiveSheet.Range("D288").Formula = sumFormula
    estRows = ActiveSheet.ListObjects(1).ListRows.count
    ActiveSheet.Range("D288").AutoFill Destination:=ActiveSheet.Range("D288:D" & 288 + estRows - 1), Type:=xlFillValues
    
    ' Mostrar mensaje de finalización
    MsgBox "Resumen completado, revisa que los estudios especiales coincidan"
End Sub

Sub ASEGURADORAS2()
    Dim sheetsArray As Variant
    Dim sheetName As String
    Dim sumFormula As String
    Dim estRows As Long
        Dim i As Long
    Dim foundSheets As Boolean

    ' Array de nombres de hojas
    sheetsArray = Array("IMSS", "PEMEX", "S. MONTERREY", "M._MOVIL", "P&G", "LAPI", "INVESCLINIC", "ORORA", "AXXA")
' Construir la fórmula SUM para la celda C8
sumFormula = "=SUM("
For i = LBound(sheetsArray) To UBound(sheetsArray)
    sheetName = sheetsArray(i)
    On Error Resume Next
   If ThisWorkbook.Sheets(sheetName) Is Nothing Then
   sumFormula = sumFormula & "0"
    ' La hoja no existe en el libro actual, por lo que no se hace nada y se pasa a la siguiente iteración
Else
         sumFormula = sumFormula & "'" & sheetName & "'!C8,"
    End If
    On Error GoTo 0
Next i
' Eliminar la última coma y cerrar el paréntesis
sumFormula = Left(sumFormula, Len(sumFormula) - 1) & ")"

' Insertar la fórmula en la celda C8 de la hoja activa
ActiveSheet.Range("C8").Formula = sumFormula

' Determinar el número de filas en la tabla ESTUDIO57 si existe
If Not IsError(ActiveSheet.ListObjects(4)) Then
    estRows = ActiveSheet.ListObjects(4).ListRows.count

    ' Rellenar el rango desde C8 hacia abajo hasta el número de filas en la tabla ESTUDIO57
    ActiveSheet.Range("C8").AutoFill Destination:=ActiveSheet.Range("C8:C" & 8 + estRows - 1), Type:=xlFillValues
End If
  
    ' Repetir el proceso para las demás tablas y celdas específicas
    ' Construir la fórmula SUM para la celda D130
    sumFormula = "=SUM("
    For i = LBound(sheetsArray) To UBound(sheetsArray)
        sheetName = sheetsArray(i)
        On Error Resume Next
        If ThisWorkbook.Sheets(sheetName) Is Nothing Then
    sumFormula = sumFormula & "0"
    ' La hoja no existe en el libro actual, por lo que no se hace nada y se pasa a la siguiente iteración
Else
            sumFormula = sumFormula & "'" & sheetName & "'!D130,"
        End If
        On Error GoTo 0
    Next i
    ' Eliminar la última coma y cerrar el paréntesis
    sumFormula = Left(sumFormula, Len(sumFormula) - 1) & ")"
    
    ' Insertar la fórmula en la celda D130 de la hoja activa
    ActiveSheet.Range("D130").Formula = sumFormula
    estRows = ActiveSheet.ListObjects(3).ListRows.count
    ActiveSheet.Range("D130").AutoFill Destination:=ActiveSheet.Range("D130:D" & 130 + estRows - 1), Type:=xlFillValues
    
    ' Construir la fórmula SUM para la celda D143
    sumFormula = "=SUM("
    For i = LBound(sheetsArray) To UBound(sheetsArray)
        sheetName = sheetsArray(i)
        On Error Resume Next
        If ThisWorkbook.Sheets(sheetName) Is Nothing Then
    sumFormula = sumFormula & "0"
    ' La hoja no existe en el libro actual, por lo que no se hace nada y se pasa a la siguiente iteración
Else
            sumFormula = sumFormula & "'" & sheetName & "'!D143,"
        End If
        On Error GoTo 0
    Next i
    sumFormula = Left(sumFormula, Len(sumFormula) - 1) & ")"
    
    ' Insertar la fórmula en la celda D143 de la hoja activa
    ActiveSheet.Range("D143").Formula = sumFormula
    estRows = ActiveSheet.ListObjects(2).ListRows.count
    ActiveSheet.Range("D143").AutoFill Destination:=ActiveSheet.Range("D143:D" & 143 + estRows - 1), Type:=xlFillValues
    
    ' Construir la fórmula SUM para la celda D288
    sumFormula = "=SUM("
    For i = LBound(sheetsArray) To UBound(sheetsArray)
        sheetName = sheetsArray(i)
        On Error Resume Next
        If ThisWorkbook.Sheets(sheetName) Is Nothing Then
    sumFormula = sumFormula & "0"
    ' La hoja no existe en el libro actual, por lo que no se hace nada y se pasa a la siguiente iteración
Else
            sumFormula = sumFormula & "'" & sheetName & "'!D288,"
        End If
        On Error GoTo 0
    Next i
    sumFormula = Left(sumFormula, Len(sumFormula) - 1) & ")"
    
    ' Insertar la fórmula en la celda D288 de la hoja activa
    ActiveSheet.Range("D288").Formula = sumFormula
    estRows = ActiveSheet.ListObjects(1).ListRows.count
    ActiveSheet.Range("D288").AutoFill Destination:=ActiveSheet.Range("D288:D" & 288 + estRows - 1), Type:=xlFillValues
    
    ' Mostrar mensaje de finalización
    MsgBox "Resumen completado, revisa que los estudios especiales coincidan"
    End Sub






