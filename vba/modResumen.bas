Attribute VB_Name = "modResumen"
' ******************************************************
' * CÓDIGO A PEGAR EN EL MÓDULO ESTÁNDAR (Ej: Modulo1) *
' ******************************************************

Option Explicit

' CONSTANTES (Alineadas con la estructura del Excel)
Private Const DATA_START_ROW As Long = 5
Private Const COL_IMPORTE As Long = 5
Private Const COL_ESTADO As Long = 8
Private Const COL_BL As Long = 3
Private Const COL_NROFACT As Long = 4
Private Const COL_DIAS_VENCIDO As Long = 7  ' Columna G (Días Vencidos - NUMÉRICO)
Private Const COL_FECHA_COBRO As Long = 10 ' Columna J

Public g_SuppressPrompts As Boolean

' --- FUNCIÓN AUXILIAR: Extrae Moneda ---
Public Function ExtractCurrencyFromFormat(ByVal nf As String) As String
    Dim re As Object, matches As Object, m As Object
    Dim s As String
    On Error GoTo Fallback
    s = Trim(nf)
    If s = "" Then ExtractCurrencyFromFormat = "": Exit Function

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "([A-Za-z]{2,4})"
    re.Global = True
    re.IgnoreCase = True

    If re.Test(s) Then
        Set matches = re.Execute(s)
        For Each m In matches
            If UCase(Trim(m.Value)) <> "AM" And UCase(Trim(m.Value)) <> "PM" Then
                ExtractCurrencyFromFormat = UCase(Trim(m.Value))
                Exit Function
            End If
        Next m
    End If

Fallback:
    If InStr(1, s, "$", vbTextCompare) > 0 Then ExtractCurrencyFromFormat = "$": Exit Function
    If InStr(1, s, "€", vbTextCompare) > 0 Then ExtractCurrencyFromFormat = "€": Exit Function
    If InStr(1, s, "£", vbTextCompare) > 0 Then ExtractCurrencyFromFormat = "£": Exit Function
    If InStr(1, s, "¥", vbTextCompare) > 0 Or InStr(1, s, "?", vbTextCompare) > 0 Then ExtractCurrencyFromFormat = "¥": Exit Function
    ExtractCurrencyFromFormat = ""
End Function


' --- FUNCIÓN AUXILIAR: Próxima Fila Disponible ---
Public Function NextAvailableRow(ByVal Sh As Worksheet, ByVal startCol As Long) As Long
    Dim lastCell As Range
    ' Definimos la primera fila de datos para Cobradas/Pendientes.
    ' La hoja FACTURACIONES usa la constante DATA_START_ROW = 5
    Dim startRowTarget As Long
    
    If UCase(Sh.Name) = "FACTURACIONES" Then
        startRowTarget = 5 ' Si por alguna razón la llamas en Facturaciones
    Else
        startRowTarget = 2 ' Para Cobradas y Pendientes, empezamos en Fila 2
    End If
    
    On Error Resume Next
    Set lastCell = Sh.Cells.Find(What:="*", After:=Sh.Cells(1, 1), LookIn:=xlFormulas, _
                      LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    
    If lastCell Is Nothing Then
        NextAvailableRow = startRowTarget ' Retorna 2 (Fila inicial para datos)
    Else
        NextAvailableRow = lastCell.Row + 1
        ' Asegura que la fila mínima sea 2
        If NextAvailableRow < startRowTarget Then NextAvailableRow = startRowTarget
    End If
End Function


' --- FUNCIÓN AUXILIAR: Verifica si ya existe en Cobradas ---
Public Function FacturaYaEnCobradas(ByVal Sh As Worksheet, ByVal invoiceNo As String, ByVal bl As String) As Boolean
    Dim f As Range, firstAddress As String
    FacturaYaEnCobradas = False
    If Trim(invoiceNo) = "" Then Exit Function
    
    With Sh.Columns(COL_NROFACT)
        Set f = .Find(What:=invoiceNo, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
        If Not f Is Nothing Then
            firstAddress = f.Address
            Do
                If Trim(CStr(Sh.Cells(f.Row, COL_BL).Value)) = bl Then
                    FacturaYaEnCobradas = True
                    Exit Function
                End If
                Set f = .FindNext(f)
                If f Is Nothing Then Exit Do
            Loop While f.Address <> firstAddress
        End If
    End With
End Function


' --- SUB: Limpia Cobradas (Borra filas con fecha de cobro > 30 días) ---
Public Sub LimpiarCobradas()
    Dim wb As Workbook
    Dim shDest As Worksheet
    Dim lastRow As Long, r As Long
    Dim datePaid As Variant
    Dim cutoffDate As Date
    
    On Error GoTo ErrHandler
    
    Set wb = ThisWorkbook
    On Error Resume Next
    Set shDest = wb.Worksheets("Cobradas")
    On Error GoTo ErrHandler
    
    If shDest Is Nothing Then Exit Sub
    
    cutoffDate = Date - 30 ' Hoy menos 30 días
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    lastRow = shDest.Cells(shDest.Rows.Count, 2).End(xlUp).Row
    
    For r = lastRow To DATA_START_ROW Step -1
        datePaid = shDest.Cells(r, COL_FECHA_COBRO).Value ' Columna J en Cobradas
        
        If IsDate(datePaid) Then
            If datePaid <= cutoffDate Then
                shDest.Rows(r).Delete Shift:=xlUp
            End If
        End If
    Next r
    
    shDest.Columns("B:K").AutoFit
    
Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    If Not g_SuppressPrompts Then MsgBox "Error en LimpiarCobradas: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume Cleanup
End Sub


' --- SUB: Genera Resumen de Pendientes ---
Public Sub GenerarResumenPendientesSimple()
    Dim wb As Workbook
    Dim shSrc As Worksheet, shDest As Worksheet
    Dim lastRow As Long, r As Long
    Dim totalPend As Double, totalVenc As Double
    Dim sum0_30 As Double, sum31_60 As Double, sum61_90 As Double, sum90p As Double ' Buckets
    Dim valEstado As String
    Dim valImporte As Variant
    Dim valDiasVencido As Variant
    Dim sampleFmt As String
    Dim daysOverdue As Double

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set wb = ThisWorkbook
    Set shSrc = wb.Worksheets("FACTURACIONES")

    On Error Resume Next
    Set shDest = wb.Worksheets("Pendientes")
    If shDest Is Nothing Then
        Set shDest = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        shDest.Name = "Pendientes"
    End If
    On Error GoTo ErrHandler

    totalPend = 0: totalVenc = 0: sum0_30 = 0: sum31_60 = 0: sum61_90 = 0: sum90p = 0
    lastRow = shSrc.Cells(shSrc.Rows.Count, 2).End(xlUp).Row
    If lastRow < DATA_START_ROW Then GoTo Finish
    
    sampleFmt = ""
    For r = DATA_START_ROW To lastRow
        If Trim(CStr(shSrc.Cells(r, COL_IMPORTE).Value)) <> "" Then
            sampleFmt = shSrc.Cells(r, COL_IMPORTE).NumberFormat
            Exit For
        End If
    Next r

    For r = DATA_START_ROW To lastRow
        valEstado = Trim(CStr(shSrc.Cells(r, COL_ESTADO).Value))
        
        If UCase(valEstado) = "PENDIENTE" Then
            valImporte = shSrc.Cells(r, COL_IMPORTE).Value
            valDiasVencido = shSrc.Cells(r, COL_DIAS_VENCIDO).Value ' LEE COLUMNA G
            
            If IsNumeric(valImporte) Then totalPend = totalPend + valImporte
            
            If IsNumeric(valDiasVencido) Then
                daysOverdue = valDiasVencido
            Else
                daysOverdue = -99999
            End If
            
            If IsNumeric(daysOverdue) And daysOverdue > 0 Then
                If IsNumeric(valImporte) Then totalVenc = totalVenc + valImporte
            End If
            
            If IsNumeric(daysOverdue) Then
                If daysOverdue <= 0 Then
                    ' A TIEMPO
                ElseIf daysOverdue <= 30 Then
                    If IsNumeric(valImporte) Then sum0_30 = sum0_30 + valImporte
                ElseIf daysOverdue <= 60 Then
                    If IsNumeric(valImporte) Then sum31_60 = sum31_60 + valImporte
                ElseIf daysOverdue <= 90 Then
                    If IsNumeric(valImporte) Then sum61_90 = sum61_90 + valImporte
                Else ' daysOverdue > 90
                    If IsNumeric(valImporte) Then sum90p = sum90p + valImporte
                End If
            End If
        End If
    Next r

    ' Escribir encabezados y totales en la hoja Pendientes (B1:G2)
    shDest.Range("B1").Value = "Total Pendiente"
    shDest.Range("C1").Value = "Total Vencido"
    shDest.Range("D1").Value = "0-30 días vencido"
    shDest.Range("E1").Value = "31-60 días vencido"
    shDest.Range("F1").Value = "61-90 días vencido"
    shDest.Range("G1").Value = "+90 días vencido"
    shDest.Range("B1:G1").Font.Bold = True

    shDest.Range("B2").Value = totalPend
    shDest.Range("C2").Value = totalVenc
    shDest.Range("D2").Value = sum0_30
    shDest.Range("E2").Value = sum31_60
    shDest.Range("F2").Value = sum61_90
    shDest.Range("G2").Value = sum90p

    If sampleFmt <> "" Then shDest.Range("B2:G2").NumberFormat = sampleFmt
    shDest.Columns("B:G").AutoFit

Finish:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    If Not g_SuppressPrompts Then MsgBox "Error en GenerarResumenPendientesSimple: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume Finish
End Sub


' --- SUB WRAPPER: Actualización Automática ---
Public Sub AutoUpdateResumen(Optional ByVal Silent As Boolean = True)
    On Error GoTo ErrHandler
    g_SuppressPrompts = Silent
    GenerarResumenPendientesSimple
    LimpiarCobradas ' Se llama después del resumen para asegurar que el resumen sea del estado actual
    g_SuppressPrompts = False
    Exit Sub
ErrHandler:
    g_SuppressPrompts = False
    If Not Silent Then
        MsgBox "Error en AutoUpdateResumen: " & Err.Number & " - " & Err.Description, vbExclamation
    End If
End Sub

