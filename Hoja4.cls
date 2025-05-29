VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CalcularValor(H As Range, AA As Range, AB As Range, k As Range)
    Dim i As Long
    Dim valor As Integer
    Dim cuenta As Integer
    Dim pantalla As Integer
    
    cuenta = 0
    pantalla = 0
    
    For i = 5 To H.Worksheet.Cells(H.Worksheet.Rows.Count, "H").End(xlUp).Row
        If H.Worksheet.Cells(i, "H").value = "" Then Exit For
        If H.Worksheet.Cells(i, "H").value = H.value Then
            If H.Worksheet.Cells(i, "AA").value = "Cuenta" Then
                cuenta = cuenta + 1
            ElseIf H.Worksheet.Cells(i, "AA").value = "Pantalla" Then
                pantalla = pantalla + 1
            End If
        End If
    Next i
    
    If cuenta > 0 Or pantalla > 0 Then
        Select Case k.value
            Case "Netflix": valor = 5
            Case "Disney": valor = 7
            Case "Star": valor = 7
            Case "Prime Video": valor = 4
            Case "Max": valor = 4
            Case "Spotify": valor = 6
            Case "YouTube", "Paramount", "Movistar TV", "IP TV", "DirecTv Go", "Crunchyroll", "Apple Tv": valor = 1
        End Select
        MsgBox "El valor calculado es: " & valor
    Else
        MsgBox "No se encontraron datos para calcular el valor."
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range

    ' Define la última fila en la columna BS (puedes ajustarla según tu necesidad)
    lastRow = Me.Cells(Me.Rows.Count, "BS").End(xlUp).Row

    ' Verifica si el cambio ocurrió en las columnas BS, BM, BD o AV
    Set rng = Intersect(Target, Me.Range("BS:BS, BM:BM, BD:BD, AV:AV"))
    If Not rng Is Nothing Then
        For Each cell In rng
            ' Verifica si el valor es una fecha válida
            If IsDate(cell.value) Then
                cell.NumberFormat = "DD/MM/YYYY"
            End If
        Next cell
        ' Recalcula todas las fórmulas en la hoja de cálculo
        Application.Calculate
    End If
End Sub

