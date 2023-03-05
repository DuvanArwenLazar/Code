Option Explicit
Sub automation()

    ' Verificacion de Rutas
    If ThisWorkbook.Sheets("Automatizacion").Range("A6").Value = "" Then
        MsgBox "La ruta de 'Al Ruedo Codificacion' NO Puede Estar Vacia"
        Exit Sub
    ElseIf ThisWorkbook.Sheets("Automatizacion").Range("A12").Value = "" Then
        MsgBox "La ruta de 'Liquidacion' NO Puede Estar Vacia"
        Exit Sub
    End If
    
    Dim column As Range
    Dim value_sought As String
    Dim al_ruedo_book As Workbook
    Dim sheets_book As Worksheet
    Dim last_row As Long
    
    Dim datahub_column As Range
    
    Dim row As Range
    Dim counter As Integer
    Dim row_value As String
    
    counter = 1
    
    Dim copied_row As Range
    Dim aux As Integer
    
    aux = 1
    
    ' --- Asignacion de valor y abrir archivo
    value_sought = "Codificaciones"
    
    Set al_ruedo_book = Workbooks.Open(Sheets("Automatizacion").Range("A6").Value)
    
    ' --- Buscar e insertar la columna necesaria
    For Each sheets_book In al_ruedo_book.Worksheets
        If sheets_book.Name = "Nacional Cacharreros" Or sheets_book.Name = "Nacional Abarroteros" Or sheets_book.Name = "Costa Abarroteros" Or sheets_book.Name = "Costa Cacharreros" Or sheets_book.Name = "Antioquia Cacharreros" Or sheets_book.Name = "Antioquia Abarrotero" Then
            Dim cell As Range
            Set column = sheets_book.Range("AG:AG")
            Set cell = column.Find(What:=value_sought, LookIn:=xlValues, LookAt:=xlWhole)
            
            If cell Is Nothing Then
                sheets_book.Columns("AE").Insert
                sheets_book.Columns.Range("AE:AE").ClearFormats
            End If
        End If
    Next sheets_book
    
    Worksheets.Add.Name = "Consolidado"
    
    Dim j As Integer
    
    ' --- Escribir la informacion de todas las hojas en una.
    For Each sheets_book In al_ruedo_book.Worksheets
        If sheets_book.Name = "Nacional Cacharreros" Or sheets_book.Name = "Nacional Abarroteros" Or sheets_book.Name = "Costa Abarroteros" Or sheets_book.Name = "Costa Cacharreros" Or sheets_book.Name = "Antioquia Cacharreros" Or sheets_book.Name = "Antioquia Abarrotero" Then
            last_row = sheets_book.Range("A" & Rows.Count).End(xlUp).row
            
            For j = 1 To last_row + 1
                ' -- Pequeños controles
                If counter = 1 Or counter = 2 Then
                    Sheets(sheets_book.Name).Rows(counter).Copy
                    Sheets("Consolidado").Range("A" & aux).PasteSpecial xlPasteAll
                    
                    counter = counter + 1
                    aux = aux + 1
                ElseIf counter > last_row Then
                    counter = 1
                ElseIf counter <= last_row And counter > 1 And counter > 2 Then
                    ' - Pasamos las filas
                    Sheets(sheets_book.Name).Rows(counter).Copy
                    Sheets("Consolidado").Range("A" & aux).PasteSpecial xlPasteAll
                    
                    ' - Rango
                    Sheets("Consolidado").Range("AO" & aux).Value = sheets_book.Name
                    
                    counter = counter + 1
                    aux = aux + 1
                End If
            Next j
        End If
    Next sheets_book
    
    ' --- Rango (Nombre de la hoja a la que pertenece)
    Columns("A").Insert
    
    Dim from_col As String
    Dim to_col As String
    Dim last_row_of_col As Long
    
    from_col = "AP"
    to_col = "A"
    
    last_row_of_col = Sheets("Consolidado").Cells(Rows.Count, from_col).End(xlUp).row
    
    ActiveSheet.Range(to_col & "1:" & to_col & last_row_of_col).Value = Sheets("Consolidado").Range(from_col & "1:" & from_col & last_row_of_col).Value

    ' --- Eliminacion de los titulos
    Dim i As Long
    Dim lastRow As Long
    lastRow = Sheets("Consolidado").Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious).row
    For i = lastRow To 1 Step -1
        If Sheets("Consolidado").Cells(i, 1).Value = "Primer Datahub" Then
            ActiveSheet.Rows(i).EntireRow.Delete
        End If
    Next i
    
    Range("A2").Value = "Rango"
    Range("A2").Interior.Color = RGB(146, 208, 60)
    Range("A2").Font.Color = RGB(255, 255, 255)
    
    Range("AP:AP").ClearContents

    ' --- Cambio de % a "Codificaciones"
    Dim l_column As Long
    Dim formula_cell, formula_completed As String
    Dim parameters() As String
    Dim l As Integer
    
    Dim cell_counter As Integer
    cell_counter = 2
    
    Sheets("Resultado").Range("L1").Value = "Codificaciones"
    
    l_column = Sheets("Resultado").Range("L" & Rows.Count).End(xlUp).row
    For l = 2 To l_column
        formula_cell = Sheets("Resultado").Range("L" & l).Formula
        
        parameters = Split(formula_cell, ",")
        formula_completed = parameters(0) & "," & parameters(1) & "," & 33 & "," & parameters(3)
        Sheets("Resultado").Range("L" & cell_counter).Formula = formula_completed
        
        Sheets("Resultado").Range("L" & l).NumberFormat = "General"
        cell_counter = cell_counter + 1
    Next l
    
    ' --- Referencia Bimestre (Buscarv + Recorrer valores)
    Dim reference_bimester As Long
    Dim reference_counter As Integer
    Dim liquidacion_book As Workbook
    Dim m As Integer
    reference_counter = 3
    
    Set liquidacion_book = Workbooks.Open(ThisWorkbook.Sheets("Automatizacion").Range("A12").Value)
    liquidacion_book.Activate
    
    reference_bimester = Sheets("Liquidación Al Ruedo ND22").Range("A" & Rows.Count).End(xlUp).row
    For m = 3 To reference_bimester
        Sheets("Liquidacion Al Ruedo ND22").Range("AE" & reference_counter).FormulaLocal = "=BUSCARV(H" & reference_counter & ",'[Al Ruedo Codificación ND.xlsx]Resultado'!$A:$L,12,FALSO)"
        reference_counter = reference_counter + 1
    Next m
End Sub
