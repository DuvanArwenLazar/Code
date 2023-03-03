Option Explicit
Sub Process_auto()

    ' Declaracion de variables
    Dim column As Range
    Dim value_sought As String
    Dim codificacion_nd_book As Workbook
    Dim sheets_book As Worksheet
    
    Dim datahub_column As Range
    
    Dim row As Range
    Dim counter As Integer
    Dim row_value As String
    
    counter = 1
    
    Dim copied_row As Range
    Dim aux As Integer
    
    aux = 1
    
    ' Asignacion de valor y abrir archivo
    value_sought = "Codificaciones"
    
    Set codificacion_nd_book = Workbooks.Open("C:\Users\duvan.espinal\Desktop\Automatizacion PYG\Al Ruedo Codificación ND.xlsx")
    
    ' Buscar de la columna
    For Each sheets_book In codificacion_nd_book.Worksheets
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
    
    ' Escribir la informacion de todas las hojas en una.
    For Each sheets_book In codificacion_nd_book.Worksheets
    
        ' Recorrer Columna De Valores
        If sheets_book.Name = "Nacional Cacharreros" Or sheets_book.Name = "Nacional Abarroteros" Or sheets_book.Name = "Costa Abarroteros" Or sheets_book.Name = "Costa Cacharreros" Or sheets_book.Name = "Antioquia Cacharreros" Or sheets_book.Name = "Antioquia Abarrotero" Then
            Set datahub_column = sheets_book.Range("A3:A250")
            For Each row In datahub_column
                If row.Value <> "" Then
                    ' Pasamos las filas
                    
                    ' Metodo 1
                    ' Set copied_row = sheets_book.Range("A" & counter & ":AZ" & counter).Value
                    ' Sheets("Duplicate").Range("A" & counter & ":AZ" & counter) = copied_row
                    
                    ' Metodo 2 (Funciono mejor en este caso)
                    Sheets(sheets_book.Name).Rows(counter).Copy
                    Sheets("Consolidado").Range("A" & aux).PasteSpecial xlPasteAll
                    
                    counter = counter + 1
                    aux = aux + 1
                Else
                    counter = 1
                End If
            Next row
        End If
    Next sheets_book
    
    ' Rango
    Columns("A").Insert
    Range("A2").Value = "Rango"
    Range("A2").Interior.Color = RGB(146, 208, 60)
    Range("A2").Font.Color = RGB(255, 255, 255)
    
    For Each sheets_book In codificacion_nd_book.Worksheets
    
        ' Recorrer Columna De Valores
        If sheets_book.Name = "Nacional Cacharreros" Or sheets_book.Name = "Nacional Abarroteros" Or sheets_book.Name = "Costa Abarroteros" Or sheets_book.Name = "Costa Cacharreros" Or sheets_book.Name = "Antioquia Cacharreros" Or sheets_book.Name = "Antioquia Abarrotero" Then
            Set datahub_column = sheets_book.Range("A3:A250")
            For Each row In datahub_column
                If row.Value <> "" Then
                    Sheets("Consolidado").Range("A" & aux).Value = "Hola"
                    counter = counter + 1
                    aux = aux + 1
                Else
                    counter = 1
                End If
            Next row
        End If
    Next sheets_book
    
    
    ' Eliminacion de los titulos (No decidido aun)
    Dim i As Long
    Dim LastRow As Long
    LastRow = Sheets("Consolidado").Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious).row
    For i = LastRow To 1 Step -1
        If Sheets("Consolidado").Cells(i, 1).Value = "Primer Datahub" Then
            ActiveSheet.Rows(i).EntireRow.Delete
        End If
    Next i

    Dim l_column As Range
    Dim formula_cell, formula_completed As String
    Dim parameters() As String
    
    Dim cell_counter As Integer
    cell_counter = 2
    
    Sheets("Resultado").Range("L1").Value = "Codificaciones"
    Set l_column = Sheets("Resultado").Range("L2:L300")
    
    For Each cell In l_column
        If cell.Value <> "" Then
            formula_cell = cell.Formula
            
            parameters = Split(formula_cell, ",")
            formula_completed = parameters(0) & "," & parameters(1) & "," & 33 & "," & parameters(3)
            Sheets("Resultado").Range("L" & cell_counter).Formula = formula_completed
            
            cell.NumberFormat = "General"
            cell_counter = cell_counter + 1
        End If
    Next cell
    
    Dim reference_bimester As Range
    Dim reference_counter As Integer
    Dim liquidacion_ruedo_book As Workbook
    reference_counter = 3
    
    Set liquidacion_ruedo_book = Workbooks.Open("C:\Users\duvan.espinal\Desktop\Automatizacion PYG\Liquidación Al Ruedo - NovDic22.xlsm")
    liquidacion_ruedo_book.Activate
    
    Set reference_bimester = Sheets("Liquidación Al Ruedo ND22").Range("AE3:AE245")
    
    For Each cell In reference_bimester
        cell.FormulaLocal = "=BUSCARV(H" & reference_counter & ",'[Al Ruedo Codificación ND.xlsx]Resultado'!$A:$L,12,FALSO)"
        reference_counter = reference_counter + 1
    Next cell
End Sub
