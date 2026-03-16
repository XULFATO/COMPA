Attribute VB_Name = "ModComparador"
Option Explicit

Private Sub ActualizarEstado(msg As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("MENU")
    ws.Range("B20").Value = "Estado: " & msg
    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = False
End Sub

Private Sub ImportarHoja(idx As Integer)
    Dim abiertos() As String
    Dim n As Integer, i As Integer
    Dim wb As Workbook, ws As Worksheet
    Dim resp As String

    n = 0
    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            ReDim Preserve abiertos(n)
            abiertos(n) = wb.Name
            n = n + 1
        End If
    Next wb

    If n = 0 Then
        MsgBox "No hay ningun Excel abierto ademas de este." & vbNewLine & _
               "Abre primero los dos archivos de datos.", vbExclamation, "Sin archivos"
        Exit Sub
    End If

    Dim lista As String
    lista = "Libros abiertos:" & vbNewLine
    For i = 0 To n - 1
        lista = lista & "  " & (i + 1) & ") " & abiertos(i) & vbNewLine
    Next i
    lista = lista & vbNewLine & "Escribe el NUMERO del libro para HOY " & idx & ":"

    resp = InputBox(lista, "Importar HOY " & idx)
    If resp = "" Then Exit Sub

    Dim sel As Integer
    sel = CInt(resp) - 1
    If sel < 0 Or sel >= n Then
        MsgBox "Numero fuera de rango.", vbExclamation
        Exit Sub
    End If

    Dim wbOrigen As Workbook
    Set wbOrigen = Workbooks(abiertos(sel))

    Dim hojas() As String
    Dim nh As Integer
    nh = 0
    For Each ws In wbOrigen.Sheets
        ReDim Preserve hojas(nh)
        hojas(nh) = ws.Name
        nh = nh + 1
    Next ws

    Dim listaH As String
    listaH = "Hojas en " & wbOrigen.Name & ":" & vbNewLine
    For i = 0 To nh - 1
        listaH = listaH & "  " & (i + 1) & ") " & hojas(i) & vbNewLine
    Next i
    listaH = listaH & vbNewLine & "Escribe el NUMERO de la hoja:"

    Dim respH As String
    respH = InputBox(listaH, "Seleccionar hoja")
    If respH = "" Then Exit Sub

    Dim selH As Integer
    selH = CInt(respH) - 1
    If selH < 0 Or selH >= nh Then
        MsgBox "Numero fuera de rango.", vbExclamation
        Exit Sub
    End If

    Dim wsOrigen As Worksheet
    Set wsOrigen = wbOrigen.Sheets(hojas(selH))

    Dim nombreBase As String
    nombreBase = Left(wbOrigen.Name, InStrRev(wbOrigen.Name, ".") - 1)
    If nombreBase = "" Then nombreBase = wbOrigen.Name
    Dim nombreDestino As String
    nombreDestino = nombreBase & " - HOY " & idx

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(nombreDestino).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    wsOrigen.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = nombreDestino

    Dim wsMenu As Worksheet
    Set wsMenu = ThisWorkbook.Sheets("MENU")
    wsMenu.Range("A" & (50 + idx)).Value = nombreDestino

    ActualizarEstado "HOY " & idx & " importado -> """ & nombreDestino & """ OK"
    MsgBox "Hoja importada correctamente:" & vbNewLine & nombreDestino, _
           vbInformation, "Importacion OK"
End Sub

Public Sub BtnImportar1()
    ImportarHoja 1
End Sub

Public Sub BtnImportar2()
    ImportarHoja 2
End Sub

Public Sub BtnComparar()
    Dim wsMenu As Worksheet
    Set wsMenu = ThisWorkbook.Sheets("MENU")

    Dim nombre1 As String, nombre2 As String
    nombre1 = wsMenu.Range("A51").Value
    nombre2 = wsMenu.Range("A52").Value

    If nombre1 = "" Or nombre2 = "" Then
        MsgBox "Primero importa las dos hojas (HOY 1 y HOY 2).", vbExclamation
        Exit Sub
    End If

    Dim ws1 As Worksheet, ws2 As Worksheet
    On Error Resume Next
    Set ws1 = ThisWorkbook.Sheets(nombre1)
    Set ws2 = ThisWorkbook.Sheets(nombre2)
    On Error GoTo 0

    If ws1 Is Nothing Or ws2 Is Nothing Then
        MsgBox "No se encuentran las hojas importadas. Vuelve a importarlas.", vbExclamation
        Exit Sub
    End If

    ActualizarEstado "Comparando..."
    Application.ScreenUpdating = False

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("COMPARACION").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Dim wsDiff As Worksheet
    Set wsDiff = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsDiff.Name = "COMPARACION"

    Dim maxRow As Long, maxCol As Long
    maxRow = WorksheetFunction.Max(ws1.UsedRange.Rows.Count, ws2.UsedRange.Rows.Count)
    maxCol = WorksheetFunction.Max(ws1.UsedRange.Columns.Count, ws2.UsedRange.Columns.Count)

    Dim c As Long
    For c = 1 To maxCol
        wsDiff.Cells(1, c).Value = ws1.Cells(1, c).Value
        wsDiff.Cells(1, c).Font.Bold = True
        wsDiff.Cells(1, c).Interior.Color = RGB(46, 117, 182)
        wsDiff.Cells(1, c).Font.Color = RGB(255, 255, 255)
    Next c

    Dim colDif As Long
    colDif = maxCol + 1
    wsDiff.Cells(1, colDif).Value = "DIFERENTE"
    wsDiff.Cells(1, colDif).Font.Bold = True
    wsDiff.Cells(1, colDif).Interior.Color = RGB(197, 90, 17)
    wsDiff.Cells(1, colDif).Font.Color = RGB(255, 255, 255)

    Dim colInfo As Long
    colInfo = maxCol + 2
    wsDiff.Cells(1, colInfo).Value = "DETALLE CAMBIOS"
    wsDiff.Cells(1, colInfo).Font.Bold = True
    wsDiff.Cells(1, colInfo).Interior.Color = RGB(197, 90, 17)
    wsDiff.Cells(1, colInfo).Font.Color = RGB(255, 255, 255)

    Dim r As Long
    Dim filaOut As Long
    filaOut = 2

    For r = 2 To maxRow
        Dim esDif As Boolean
        esDif = False
        Dim detalle As String
        detalle = ""

        For c = 1 To maxCol
            Dim v1 As Variant, v2 As Variant
            v1 = ws1.Cells(r, c).Value
            v2 = ws2.Cells(r, c).Value
            If CStr(v1) <> CStr(v2) Then
                esDif = True
                detalle = detalle & ws1.Cells(1, c).Value & ": [" & v1 & "] -> [" & v2 & "]  |  "
            End If
        Next c

        For c = 1 To maxCol
            wsDiff.Cells(filaOut, c).Value = ws1.Cells(r, c).Value
        Next c

        If esDif Then
            wsDiff.Cells(filaOut, colDif).Value = "SI"
            wsDiff.Cells(filaOut, colDif).Font.Color = RGB(192, 0, 0)
            wsDiff.Cells(filaOut, colDif).Font.Bold = True
            wsDiff.Rows(filaOut).Interior.Color = RGB(255, 235, 230)
            If Len(detalle) > 2 Then
                wsDiff.Cells(filaOut, colInfo).Value = Left(detalle, Len(detalle) - 3)
            End If
        Else
            wsDiff.Cells(filaOut, colDif).Value = "NO"
            wsDiff.Cells(filaOut, colDif).Font.Color = RGB(55, 86, 35)
        End If

        filaOut = filaOut + 1
    Next r

    wsDiff.Columns.AutoFit
    wsDiff.Range(wsDiff.Cells(1, 1), wsDiff.Cells(1, colInfo)).AutoFilter

    wsDiff.Activate
    wsDiff.Rows(2).Select
    ActiveWindow.FreezePanes = True
    wsDiff.Range("A1").Select

    Dim numDif As Long
    numDif = 0
    Dim cell As Range
    For Each cell In wsDiff.Range(wsDiff.Cells(2, colDif), wsDiff.Cells(filaOut - 1, colDif))
        If cell.Value = "SI" Then numDif = numDif + 1
    Next cell

    Application.ScreenUpdating = True
    wsMenu.Activate

    ActualizarEstado "Comparacion lista · " & numDif & " filas diferentes de " & (filaOut - 2) & " totales"
    MsgBox "Comparacion completada." & vbNewLine & vbNewLine & _
           "Filas diferentes: " & numDif & vbNewLine & _
           "Filas totales: " & (filaOut - 2) & vbNewLine & vbNewLine & _
           "Hoja COMPARACION creada con autofiltro activado.", _
           vbInformation, "Resultado"

    wsDiff.Activate
End Sub

Public Sub CrearBotones()
    Dim wsMenu As Worksheet
    Set wsMenu = ThisWorkbook.Sheets("MENU")

    Dim shp As Shape
    For Each shp In wsMenu.Shapes
        shp.Delete
    Next shp

    Dim btn1 As Shape
    Set btn1 = wsMenu.Shapes.AddShape(msoShapeRoundedRectangle, _
        wsMenu.Range("C9").Left, wsMenu.Range("C9").Top, _
        wsMenu.Range("F9").Left + wsMenu.Range("F9").Width - wsMenu.Range("C9").Left, _
        wsMenu.Range("C11").Top + wsMenu.Range("C11").Height - wsMenu.Range("C9").Top)
    With btn1
        .Name = "BtnHoy1"
        .TextFrame.Characters.Text = "IMPORTAR HOY 1"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 11
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .Fill.ForeColor.RGB = RGB(46, 117, 182)
        .Line.Visible = msoFalse
        .OnAction = "BtnImportar1"
    End With

    Dim btn2 As Shape
    Set btn2 = wsMenu.Shapes.AddShape(msoShapeRoundedRectangle, _
        wsMenu.Range("G9").Left, wsMenu.Range("G9").Top, _
        wsMenu.Range("I9").Left + wsMenu.Range("I9").Width - wsMenu.Range("G9").Left, _
        wsMenu.Range("G11").Top + wsMenu.Range("G11").Height - wsMenu.Range("G9").Top)
    With btn2
        .Name = "BtnHoy2"
        .TextFrame.Characters.Text = "IMPORTAR HOY 2"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 11
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .Fill.ForeColor.RGB = RGB(46, 117, 182)
        .Line.Visible = msoFalse
        .OnAction = "BtnImportar2"
    End With

    Dim btn3 As Shape
    Set btn3 = wsMenu.Shapes.AddShape(msoShapeRoundedRectangle, _
        wsMenu.Range("D16").Left, wsMenu.Range("D16").Top, _
        wsMenu.Range("H16").Left + wsMenu.Range("H16").Width - wsMenu.Range("D16").Left, _
        wsMenu.Range("D18").Top + wsMenu.Range("D18").Height - wsMenu.Range("D16").Top)
    With btn3
        .Name = "BtnComparar"
        .TextFrame.Characters.Text = "COMPARAR Y GENERAR INFORME"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 11
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .Fill.ForeColor.RGB = RGB(55, 86, 35)
        .Line.Visible = msoFalse
        .OnAction = "BtnComparar"
    End With
End Sub
