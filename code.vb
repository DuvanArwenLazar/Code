Option Explicit
Sub Process_auto()

    ' Declaracion de variables
    Dim column As Range
    Dim value_sought As String
    Dim search_book As Workbook
    Dim sheets_book As Worksheet
    
    Dim datahub_column As Range
    
    Dim row As Range
    Dim counter As String
    Dim row_value As String
    
    counter = 1
    
    Dim copied_row As Range
    Dim aux As Integer
    
    aux = 1
    
    ' Asignacion de valor y abrir archivo
    value_sought = "Codificaciones"
    
    Set search_book = Workbooks.Open("C:\Users\duvan.espinal\Desktop\Automatizacion PYG\Al Ruedo Codificación ND.xlsx")
    
    ' Buscar de la columna
    For Each sheets_book In search_book.Worksheets
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
    
    Worksheets.Add.Name = "Duplicate"
    
    ' Escribir la informacion de todas las hojas en una.
    For Each sheets_book In search_book.Worksheets
    
        ' Recorrer Columna De Valores
        If sheets_book.Name = "Nacional Cacharreros" Or sheets_book.Name = "Nacional Abarroteros" Or sheets_book.Name = "Costa Abarroteros" Or sheets_book.Name = "Costa Cacharreros" Or sheets_book.Name = "Antioquia Cacharreros" Or sheets_book.Name = "Antioquia Abarrotero" Then
            Set datahub_column = sheets_book.Range("A3:A150")
            For Each row In datahub_column
                If row.Value <> "" Then
                    ' Pasamos las filas
                    
                    ' Metodo 1
                    ' Set copied_row = sheets_book.Range("A" & counter & ":AZ" & counter).Value
                    ' Sheets("Duplicate").Range("A" & counter & ":AZ" & counter) = copied_row
                    
                    ' Metodo 2 (Funciono mejor en este caso)
                    Sheets(sheets_book.Name).Rows(counter).Copy
                    Sheets("Duplicate").Range("A" & aux).PasteSpecial xlPasteAll
                    
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
    LastRow = Sheets("Duplicate").Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious).row
    For i = LastRow To 1 Step -1
        If Sheets("Duplicate").Cells(i, 1).Value = "Primer Datahub" Then
            ActiveSheet.Rows(i).EntireRow.Delete
        End If
    Next i
End Sub