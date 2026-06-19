Attribute VB_Name = "Hoja_FACTURACIONES"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
' ***************************************************************
' * CÓDIGO A PEGAR EN EL MÓDULO DE LA HOJA "FACTURACIONES" *
' ***************************************************************

Option Explicit

' CONSTANTES DE COLUMNAS (Repetidas para claridad, pero usan COL_ESTADO=8, COL_FECHA_COBRO=10)
Private Const DATA_START_ROW As Long = 5
Private Const SRC_COL_START As Long = 2
Private Const SRC_COL_END As Long = 10
Private Const COL_IMPORTE As Long = 5
Private Const COL_BL As Long = 3
Private Const COL_NROFACT As Long = 4
Private Const COL_ESTADO As Long = 8
Private Const COL_FECHA_COBRO As Long = 10
Private Const COL_DIAS_VENCIDO As Long = 7 ' Columna G

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim wb As Workbook
    Dim shDest As Worksheet
    Dim inter As Range, cell As Range
    Dim rowsToProcess() As Long
    Dim rCount As Long, i As Long
    Dim srcRow As Long
    Dim invoiceNo As String, bl As String
    Dim isCharged As Boolean
    
    On Error GoTo ErrHandler
    Set wb = ThisWorkbook
    
    ' 1. DETECCIÓN DE CAMBIOS
    Set inter = Intersect(Target, Me.Columns(COL_ESTADO)) ' Solo reaccionar si cambió la columna Estado (H)
    If inter Is Nothing Then
        ' Si no es la columna H, chequea si es B2 (Fecha Actual) o G (Días Vencidos)
        Set inter = Intersect(Target, Union(Me.Range("B2"), Me.Columns(COL_DIAS_VENCIDO)))
        If Not inter Is Nothing Then
            Application.ScreenUpdating = False
            AutoUpdateResumen True
            Application.ScreenUpdating = True
        End If
        Exit Sub
    End If
    
    ' 2. Verificamos sólo celdas desde DATA_START_ROW
    rCount = 0
    For Each cell In inter.Cells
        If cell.Row >= DATA_START_ROW Then
            rCount = rCount + 1
            ReDim Preserve rowsToProcess(1 To rCount)
            rowsToProcess(rCount) = cell.Row
        End If
    Next cell
    
    If rCount = 0 Then Exit Sub
    
    ' 3. Aseguramos que exista hoja destino 'Cobradas'
    On Error Resume Next
    Set shDest = wb.Worksheets("Cobradas")
    On Error GoTo ErrHandler
    If shDest Is Nothing Then
        MsgBox "No se encuentra la hoja 'Cobradas'. Creala y pon encabezados.", vbExclamation
        Exit Sub
    End If
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' 4. PROCESAMIENTO DE FILAS CAMBIADAS
    For i = 1 To UBound(rowsToProcess)
        srcRow = rowsToProcess(i)
        isCharged = (UCase(Trim(CStr(Me.Cells(srcRow, COL_ESTADO).Value))) = "COBRADO")
        
        If isCharged Then
            
            ' A. Registrar fecha de cobro en columna J (si está vacía)
            If Not IsDate(Me.Cells(srcRow, COL_FECHA_COBRO).Value) Then
                Me.Cells(srcRow, COL_FECHA_COBRO).Value = Date
            End If
            
            ' B. Mover a Cobradas
            invoiceNo = Trim(CStr(Me.Cells(srcRow, COL_NROFACT).Value))
            bl = Trim(CStr(Me.Cells(srcRow, COL_BL).Value))
            
            If Not FacturaYaEnCobradas(shDest, invoiceNo, bl) Then
                
                Dim destRow As Long, copyRange As Range, amtFmt As String, currencyCode As String
                destRow = NextAvailableRow(shDest, SRC_COL_START)
                Set copyRange = Me.Range(Me.Cells(srcRow, SRC_COL_START), Me.Cells(srcRow, SRC_COL_END))
                
                shDest.Range(shDest.Cells(destRow, SRC_COL_START), shDest.Cells(destRow, SRC_COL_END)).Value = copyRange.Value
                
                amtFmt = Me.Cells(srcRow, COL_IMPORTE).NumberFormat
                shDest.Cells(destRow, COL_IMPORTE).NumberFormat = amtFmt
                currencyCode = ExtractCurrencyFromFormat(amtFmt)
                shDest.Cells(destRow, 11).Value = currencyCode ' Columna K para moneda
                
                copyRange.Copy
                shDest.Cells(destRow, SRC_COL_START).PasteSpecial xlPasteFormats
                Application.CutCopyMode = False
            End If
        Else
            ' Si se cambia de COBRADO a otro estado, limpia la fecha de cobro
            If IsDate(Me.Cells(srcRow, COL_FECHA_COBRO).Value) Then
                Me.Cells(srcRow, COL_FECHA_COBRO).ClearContents
            End If
        End If
    Next i
    
Cleanup:
    AutoUpdateResumen True ' Llama al wrapper para limpiar y actualizar resumen
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error en Worksheet_Change (FACTURACIONES): " & Err.Number & " - " & Err.Description, vbExclamation
    Resume Cleanup
End Sub

Private Sub Worksheet_Calculate()
    ' Actualizar en silencio cuando la hoja se recalcule
    On Error Resume Next
    AutoUpdateResumen True
End Sub
