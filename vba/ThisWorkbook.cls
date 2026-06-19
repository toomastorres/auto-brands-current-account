Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
' ********************************************************
' * CÓDIGO A PEGAR EN EL MÓDULO "ThisWorkbook" *
' ********************************************************

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ErrHandler

    If Sh Is Nothing Then Exit Sub
    If UCase(Sh.Name) <> "FACTURACIONES" Then Exit Sub

    Application.EnableEvents = False

    Dim bRun As Boolean
    bRun = False

    ' Chequear cambios en columnas relevantes (E=Importe, G=Dias, H=Estado, I=Categoría, J=Fecha Cobro)
    If Not Intersect(Target, Sh.Range("B2")) Is Nothing Then bRun = True
    If Not Intersect(Target, Sh.Range("E:E")) Is Nothing Then bRun = True
    If Not Intersect(Target, Sh.Range("G:G")) Is Nothing Then bRun = True
    If Not Intersect(Target, Sh.Range("H:H")) Is Nothing Then bRun = True
    If Not Intersect(Target, Sh.Range("I:I")) Is Nothing Then bRun = True
    If Not Intersect(Target, Sh.Range("J:J")) Is Nothing Then bRun = True

    If bRun Then
        Dim inter As Range
        Set inter = Intersect(Target, Sh.Range("B2,E5:E" & Sh.Rows.Count & ",G5:J" & Sh.Rows.Count))
        If Not inter Is Nothing Then
            AutoUpdateResumen True
        End If
    End If

ExitHandler:
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Resume ExitHandler
End Sub
