' Corrección puntual para el módulo VBA del libro informe general.xlsm
' Reemplaza el procedimiento AjustarTabla existente por este bloque.
' No importes este archivo como un segundo módulo completo: hacerlo generaría
' nombres duplicados. Guarda una copia de seguridad del libro antes de probar.
'
' Cambios incluidos:
' - Redimensiona ListObject sin borrar filas completas de la hoja, evitando
'   desplazar o eliminar celdas combinadas fuera de la tabla.
' - No usa FillDown sobre toda la tabla, pues eso puede duplicar datos fuente.
' - Solo rellena hacia abajo columnas cuya primera celda de datos contiene fórmula.
'
Private Sub AjustarTabla(ByVal n As Long)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim rangoNuevo As Range
    Dim i As Long
    Dim filas As Long
    Dim columnas As Long

    If n < 1 Then
        Err.Raise vbObjectError + 522, "AjustarTabla", _
                  "La cantidad de registros debe ser mayor que cero."
    End If

    Set ws = ThisWorkbook.Worksheets(HOJA_TABLA)

    On Error Resume Next
    Set lo = ws.ListObjects(NOMBRE_TABLA)
    On Error GoTo 0

    If lo Is Nothing Then
        Err.Raise vbObjectError + 521, "AjustarTabla", _
                  "No se encontró la tabla '" & NOMBRE_TABLA & _
                  "' en la hoja '" & HOJA_TABLA & "'."
    End If

    On Error Resume Next
    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    On Error GoTo 0

    filas = n + 1  'filas de datos más encabezado
    columnas = lo.ListColumns.Count
    Set rangoNuevo = lo.HeaderRowRange.Cells(1, 1).Resize(filas, columnas)

    ' Resize no elimina ni desplaza filas completas de la hoja.
    lo.Resize rangoNuevo

    ' Recalcula solo columnas calculadas; no sobrescribe las columnas de datos.
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListColumns.Count
            If lo.ListColumns(i).DataBodyRange.Cells(1, 1).HasFormula Then
                lo.ListColumns(i).DataBodyRange.FillDown
            End If
        Next i
    End If
End Sub
