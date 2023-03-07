Option Explicit
Sub automation()

    ' --- Verificacion de Rutas
    If ThisWorkbook.Sheets("Automatizacion").range("A6").Value = "" Then
        MsgBox "La ruta de 'Al Ruedo Codificacion' NO Puede Estar Vacia"
        Exit Sub
    ElseIf ThisWorkbook.Sheets("Automatizacion").range("A12").Value = "" Then
        MsgBox "La ruta de 'Liquidacion' NO Puede Estar Vacia"
        Exit Sub
    End If
    
    Dim column As range
    Dim value_sought As String
    Dim al_ruedo_book As Workbook
    Dim sheets_book As Worksheet
    Dim last_row As Long
    
    Dim datahub_column As range
    
    Dim row As range
    Dim counter As Integer
    Dim row_value As String
    
    counter = 1
    
    Dim copied_row As range
    Dim aux As Integer
    
    aux = 1
    
    ' --- Asignacion de valor y abrir archivo
    value_sought = "Codificaciones"
    
    Set al_ruedo_book = Workbooks.Open(ThisWorkbook.Sheets("Automatizacion").range("A6").Value)
    
    ' --- Buscar e insertar la columna necesaria
    For Each sheets_book In al_ruedo_book.Worksheets
        If sheets_book.Name = "Nacional Cacharreros" Or sheets_book.Name = "Nacional Abarroteros" Or sheets_book.Name = "Costa Abarroteros" Or sheets_book.Name = "Costa Cacharreros" Or sheets_book.Name = "Antioquia Cacharreros" Or sheets_book.Name = "Antioquia Abarrotero" Then
            Dim cell As range
            Set column = sheets_book.range("AG:AG")
            Set cell = column.Find(What:=value_sought, LookIn:=xlValues, LookAt:=xlWhole)
            
            If cell Is Nothing Then
                sheets_book.columns("AE").Insert
                sheets_book.columns.range("AE:AE").ClearFormats
            End If
        End If
    Next sheets_book
    
    Worksheets.Add.Name = "Consolidado"
    
    Dim j As Integer
    
    ' --- Escribir la informacion de todas las hojas en una.
    For Each sheets_book In al_ruedo_book.Worksheets
        If sheets_book.Name = "Nacional Cacharreros" Or sheets_book.Name = "Nacional Abarroteros" Or sheets_book.Name = "Costa Abarroteros" Or sheets_book.Name = "Costa Cacharreros" Or sheets_book.Name = "Antioquia Cacharreros" Or sheets_book.Name = "Antioquia Abarrotero" Then
            last_row = sheets_book.range("A" & Rows.Count).End(xlUp).row
            
            For j = 1 To last_row + 1
                ' -- Pequeños controles
                If counter = 1 Or counter = 2 Then
                    Sheets(sheets_book.Name).Rows(counter).Copy
                    Sheets("Consolidado").range("A" & aux).PasteSpecial xlPasteAll
                    
                    counter = counter + 1
                    aux = aux + 1
                ElseIf counter > last_row Then
                    counter = 1
                ElseIf counter <= last_row And counter > 1 And counter > 2 Then
                    ' - Pasamos las filas
                    Sheets(sheets_book.Name).Rows(counter).Copy
                    Sheets("Consolidado").range("A" & aux).PasteSpecial xlPasteAll
                    
                    ' - Rango
                    Sheets("Consolidado").range("AO" & aux).Value = sheets_book.Name
                    
                    counter = counter + 1
                    aux = aux + 1
                End If
            Next j
        End If
    Next sheets_book
    
    ' --- Rango (Nombre de la hoja a la que pertenece)
    columns("A").Insert
    
    Dim from_col As String
    Dim to_col As String
    Dim last_row_of_col As Long
    
    from_col = "AP"
    to_col = "A"
    
    last_row_of_col = Sheets("Consolidado").Cells(Rows.Count, from_col).End(xlUp).row
    
    ActiveSheet.range(to_col & "1:" & to_col & last_row_of_col).Value = Sheets("Consolidado").range(from_col & "1:" & from_col & last_row_of_col).Value

    ' --- Modificacion de los titulos
    Dim value_to_look As String
    Dim range As range
    Dim cell_delete As range
    Dim c As Integer
    
    value_to_look = "Primer Datahub"
    Set range = ActiveSheet.UsedRange
    
    For Each cell_delete In range
        If cell_delete.Value = value_to_look Then
            cell_delete.EntireRow.Delete
        End If
    Next cell_delete
    
    Dim column_values As Variant
    column_values = Array("", "Rango", "Primer Datahub", "Distribuidor", "Unidad", "Grupo", "UM", "Rep", "Source Store ID", "Razón social", "NIT", "Tipo FM", "Nombre Comercial")
    ' Codificaciones, Contrato, %
    For c = 1 To 12
        Cells(1, c).Value = column_values(c)
        Cells(1, c).Font.Color = RGB(255, 255, 255)
        Cells(1, c).Interior.Color = RGB(146, 208, 60)
        Cells(1, c).Borders.LineStyle = xlContinuous
        Cells(1, c).Borders.Weight = xlThin
        Cells(1, c).Borders.ColorIndex = xlAutomatic
    Next c
    
    Sheets("Consolidado").range("AP:AP").ClearContents
    
    Sheets("Consolidado").range("AH1:AJ1").Font.Color = RGB(255, 255, 255)
    Sheets("Consolidado").range("AH1:AJ1").Interior.Color = RGB(146, 208, 60)
    Sheets("Consolidado").range("AH1:AJ1").Borders.LineStyle = xlContinuous
    Sheets("Consolidado").range("AH1:AJ1").Borders.Weight = xlThin
    Sheets("Consolidado").range("AH1:AJ1").Borders.ColorIndex = xlAutomatic
    
    Sheets("Consolidado").range("AH1").Value = "Codificaciones"
    Sheets("Consolidado").range("AI1").Value = "Contrato"
    Sheets("Consolidado").range("AJ1").Value = "%"

    ' --- Cambio de % a "Codificaciones"
    Dim l_column As Long
    Dim formula_cell, formula_completed As String
    Dim parameters() As String
    Dim l As Integer
    
    Dim cell_counter As Integer
    cell_counter = 2
    
    Sheets("Resultado").range("L1").Value = "Codificaciones"
    
    l_column = Sheets("Resultado").range("L" & Rows.Count).End(xlUp).row
    For l = 2 To l_column
        formula_cell = Sheets("Resultado").range("L" & l).Formula
        
        parameters = Split(formula_cell, ",")
        formula_completed = parameters(0) & "," & parameters(1) & "," & 33 & "," & parameters(3)
        Sheets("Resultado").range("L" & cell_counter).Formula = formula_completed
        
        Sheets("Resultado").range("L" & l).NumberFormat = "General"
        cell_counter = cell_counter + 1
    Next l
    
    ' --- Referencia Bimestre (Buscarv + Recorrer valores)
    Dim reference_bimester As Long
    Dim reference_counter As Integer
    Dim liquidacion_book As Workbook
    Dim m As Integer

    Set liquidacion_book = Workbooks.Open(ThisWorkbook.Sheets("Automatizacion").range("A12").Value)
    liquidacion_book.Activate
    
    reference_bimester = Sheets("Liquidación Al Ruedo ND22").range("A" & Rows.Count).End(xlUp).row
    
    For m = 3 To reference_bimester
        liquidacion_book.Sheets("Liquidación Al Ruedo ND22").range("AE" & m).FormulaLocal = "=BUSCARV(H" & m & ";'[" & al_ruedo_book.Name & "]Resultado'!$A:$L;12;FALSO)"
        ' liquidacion_book.Sheets("Liquidación Al Ruedo ND22").range("AE" & m).FormulaLocal = "=BUSCARV(H3;'[Al Ruedo Codificación ND.xlsx]Resultado'!$A:$L;12;0)"
    Next m
End Sub

Sub total_purchases()
    ' --- Verificacion de Rutas
    If ThisWorkbook.Sheets("Automatizacion").range("A9").Value = "" Then
        MsgBox "La ruta de 'Ventas Total' NO Puede Estar Vacia"
        Exit Sub
    End If
    
    ' --- Recorrer compras totales y modificar celdas de otro archivo
    Dim liquidacion_book As Workbook
    Dim compras_total_book As Workbook
    
    Set compras_total_book = Workbooks.Open(ThisWorkbook.Sheets("Automatizacion").range("A9").Value)
    Set liquidacion_book = Workbooks.Open(ThisWorkbook.Sheets("Automatizacion").range("A12").Value)
    
    liquidacion_book.Activate
    
    Dim last_row As Long
    Dim i As Integer
    last_row = Cells(Rows.Count, "A").End(xlUp).row

    For i = 3 To last_row
        Cells(i, "AO").FormulaLocal = "=BUSCARV(H" & i & ";'[" & compras_total_book.Name & "]Sell out'!$B:$E; 3; FALSO)"
        Cells(i, "AX").FormulaLocal = "=BUSCARV(H" & i & ";'[" & compras_total_book.Name & "]Sell out'!$B:$E; 4; FALSO)"
        
        Cells(i, "AP").FormulaLocal = "=BUSCARV(I" & i & ";'[" & compras_total_book.Name & "]Sell out'!$B:$E; 3; FALSO)"
        Cells(i, "AY").FormulaLocal = "=BUSCARV(I" & i & ";'[" & compras_total_book.Name & "]Sell out'!$B:$E; 4; FALSO)"
        
        Cells(i, "AQ").FormulaLocal = "=BUSCARV(J" & i & ";'[" & compras_total_book.Name & "]Sell out'!$B:$E; 3; FALSO)"
        Cells(i, "AZ").FormulaLocal = "=BUSCARV(J" & i & ";'[" & compras_total_book.Name & "]Sell out'!$B:$E; 4; FALSO)"
        
        Cells(i, "AR").FormulaLocal = "=BUSCARV(K" & i & ";'[" & compras_total_book.Name & "]Sell out'!$B:$E; 3; FALSO)"
        Cells(i, "BA").FormulaLocal = "=BUSCARV(K" & i & ";'[" & compras_total_book.Name & "]Sell out'!$B:$E; 4; FALSO)"
        
        Cells(i, "AS").FormulaLocal = "=BUSCARV(L" & i & ";'[" & compras_total_book.Name & "]Sell out'!$B:$E; 3; FALSO)"
        Cells(i, "BB").FormulaLocal = "=BUSCARV(L" & i & ";'[" & compras_total_book.Name & "]Sell out'!$B:$E; 4; FALSO)"
        
        Cells(i, "AT").FormulaLocal = "=BUSCARV(M" & i & ";'[" & compras_total_book.Name & "]Sell out'!$B:$E; 3; FALSO)"
        Cells(i, "BC").FormulaLocal = "=BUSCARV(M" & i & ";'[" & compras_total_book.Name & "]Sell out'!$B:$E; 4; FALSO)"
        
        Cells(i, "AU").FormulaLocal = "=BUSCARV(N" & i & ";'[" & compras_total_book.Name & "]Sell out'!$B:$E; 3; FALSO)"
        Cells(i, "BD").FormulaLocal = "=BUSCARV(N" & i & ";'[" & compras_total_book.Name & "]Sell out'!$B:$E; 4; FALSO)"
        
        Cells(i, "AV").FormulaLocal = "=BUSCARV(O" & i & ";'[" & compras_total_book.Name & "]Sell out'!$B:$E; 3; FALSO)"
        Cells(i, "BE").FormulaLocal = "=BUSCARV(O" & i & ";'[" & compras_total_book.Name & "]Sell out'!$B:$E; 4; FALSO)"
        
        Cells(i, "AW").FormulaLocal = "=BUSCARV(P" & i & ";'[" & compras_total_book.Name & "]Sell out'!$B:$E; 3; FALSO)"
        Cells(i, "BF").FormulaLocal = "=BUSCARV(P" & i & ";'[" & compras_total_book.Name & "]Sell out'!$B:$E; 4; FALSO)"
    Next i
    
    Dim cleaning As range
    
    For Each cleaning In liquidacion_book.Sheets("Liquidación Al Ruedo ND22").range("AO:BF")
        If IsError(cleaning.Value) Then
            If cleaning.Value = CVErr(xlErrNA) Then
                cleaning.Value = ""
            End If
        End If
    Next cleaning
End Sub