'===========================================================
' COMPARADOR DE EXCELS - pega este bloque en un modulo VBA
' Alt+F11 -> Insertar -> Modulo -> borra lo que haya -> pega
'===========================================================

Sub ImportarHoy1()
    ImportarHoja 1
End Sub

Sub ImportarHoy2()
    ImportarHoja 2
End Sub

'-----------------------------------------------------------
' BORRAR todas las hojas excepto MENU
'-----------------------------------------------------------
Sub BorrarTodo()

    Dim resp As Integer
    resp = MsgBox("Se eliminaran todas las hojas excepto MENU." & vbCrLf & _
                  "Esto incluye las importadas y la comparacion." & vbCrLf & vbCrLf & _
                  "Continuar?", vbQuestion + vbYesNo, "Confirmar borrado")
    If resp = vbNo Then Exit Sub

    Dim wsMenu As Worksheet
    Set wsMenu = ThisWorkbook.Worksheets("MENU")
    wsMenu.Range("J1").Value = ""
    wsMenu.Range("J2").Value = ""

    Application.DisplayAlerts = False
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "MENU" Then ws.Delete
    Next ws
    Application.DisplayAlerts = True

    wsMenu.Activate
    MsgBox "Listo. Solo queda la hoja MENU.", vbInformation, "Borrado completado"
End Sub


'-----------------------------------------------------------
' IMPORTAR HOJA (slot = 1 o 2)
'-----------------------------------------------------------
Sub ImportarHoja(slot As Integer)

    Dim wsMenu As Worksheet
    Set wsMenu = ThisWorkbook.Worksheets("MENU")

    ' --- 1) Listar workbooks abiertos (excluir este) ---
    Dim nombres() As String
    Dim wbLoop    As Workbook
    Dim n         As Integer
    Dim i         As Integer
    n = 0

    For Each wbLoop In Application.Workbooks
        If wbLoop.Name <> ThisWorkbook.Name Then
            ReDim Preserve nombres(n)
            nombres(n) = wbLoop.Name
            n = n + 1
        End If
    Next wbLoop

    If n = 0 Then
        MsgBox "No hay otros ficheros Excel abiertos." & vbCrLf & _
               "Abre primero el fichero que quieres importar.", _
               vbExclamation, "Sin ficheros"
        Exit Sub
    End If

    ' --- 2) Elegir fichero ---
    Dim lista As String
    lista = "Ficheros Excel abiertos:" & vbCrLf & vbCrLf
    For i = 0 To n - 1
        lista = lista & "  " & (i + 1) & "  ->  " & nombres(i) & vbCrLf
    Next i
    lista = lista & vbCrLf & "Escribe el numero del fichero:"

    Dim respWB As Variant
    respWB = Application.InputBox(lista, "Importar HOY " & slot, Type:=1)
    If VarType(respWB) = vbBoolean Then Exit Sub

    Dim idxWB As Integer
    idxWB = CInt(respWB) - 1
    If idxWB < 0 Or idxWB >= n Then
        MsgBox "Numero fuera de rango (1 a " & n & ").", vbExclamation
        Exit Sub
    End If

    Dim wbOrigen As Workbook
    Set wbOrigen = Application.Workbooks(nombres(idxWB))

    ' --- 3) Coger siempre la primera hoja (PAGE 1) ---
    Dim wsOrigen As Worksheet
    Set wsOrigen = wbOrigen.Worksheets(1)

    ' --- 4) Nombre: nombre original + sufijo v1 o v2 ---
    Dim nomBase As String
    nomBase = wsOrigen.Name
    If Len(nomBase) > 25 Then nomBase = Left(nomBase, 25)

    Dim nomHoja As String
    nomHoja = nomBase & " v" & slot

    Dim cars As Variant
    Dim c    As Variant
    cars = Array("/", "\", "?", "*", "[", "]", ":")
    For Each c In cars
        nomHoja = Replace(nomHoja, CStr(c), "_")
    Next c

    ' --- 5) Borrar hoja anterior del mismo slot si existe ---
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(nomHoja).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' --- 6) Copiar hoja al final ---
    Dim wbThis     As Workbook
    Dim totalHojas As Integer
    Set wbThis = ThisWorkbook
    totalHojas = wbThis.Worksheets.Count
    wsOrigen.Copy After:=wbThis.Worksheets(totalHojas)
    wbThis.Worksheets(wbThis.Worksheets.Count).Name = nomHoja

    ' --- 7) Guardar referencia en MENU ---
    If slot = 1 Then
        wsMenu.Range("J1").Value = nomHoja
    Else
        wsMenu.Range("J2").Value = nomHoja
    End If

    wsMenu.Activate
    MsgBox "Hoja importada como:" & vbCrLf & vbCrLf & _
           "  << " & nomHoja & " >>", vbInformation, "HOY " & slot & " OK"
End Sub


'===========================================================
' COMPARAR
' - Cruza por Employee ID (no por posicion de fila)
' - IDs solo en v1 -> v2 en blanco, DIFERENTE = "SOLO EN V1"
' - IDs solo en v2 -> v1 en blanco, DIFERENTE = "SOLO EN V2"
' - IDs en ambos   -> compara campo a campo
' - Borde grueso al final de cada grupo de columnas
' - Ancho autoajustado sin recortar texto
'===========================================================
Sub CompararHojas()

    Dim wsMenu As Worksheet
    Set wsMenu = ThisWorkbook.Worksheets("MENU")

    Dim nomH1 As String
    Dim nomH2 As String
    nomH1 = Trim(wsMenu.Range("J1").Value)
    nomH2 = Trim(wsMenu.Range("J2").Value)

    If nomH1 = "" Or nomH2 = "" Then
        MsgBox "Importa primero las dos hojas (HOY 1 y HOY 2).", _
               vbExclamation, "Faltan hojas"
        Exit Sub
    End If

    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets(nomH1)
    Set ws2 = ThisWorkbook.Worksheets(nomH2)
    On Error GoTo 0

    If ws1 Is Nothing Then
        MsgBox "No encuentro la hoja HOY 1:" & vbCrLf & nomH1, vbCritical
        Exit Sub
    End If
    If ws2 Is Nothing Then
        MsgBox "No encuentro la hoja HOY 2:" & vbCrLf & nomH2, vbCritical
        Exit Sub
    End If

    ' --- Limites ---
    Dim lastRow1 As Long, lastCol1 As Long
    Dim lastRow2 As Long, lastCol2 As Long
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastCol1 = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    lastCol2 = ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column
    Dim maxCol As Long
    maxCol = Application.Max(lastCol1, lastCol2)

    ' --- Localizar columna Employee ID en ambas hojas ---
    Dim colEmp1 As Long
    Dim colEmp2 As Long
    Dim col     As Long
    colEmp1 = 0
    colEmp2 = 0

    For col = 1 To lastCol1
        If LCase(Trim(CStr(ws1.Cells(1, col).Value))) = "employee id" Then
            colEmp1 = col
            Exit For
        End If
    Next col
    For col = 1 To lastCol2
        If LCase(Trim(CStr(ws2.Cells(1, col).Value))) = "employee id" Then
            colEmp2 = col
            Exit For
        End If
    Next col

    If colEmp1 = 0 Or colEmp2 = 0 Then
        MsgBox "No se encuentra 'Employee ID' en " & _
               IIf(colEmp1 = 0, nomH1, nomH2) & "." & vbCrLf & _
               "Verifica que el nombre es exactamente 'Employee ID'.", _
               vbExclamation, "Columna no encontrada"
        Exit Sub
    End If

    ' --- Indice de ws2: Employee ID -> numero de fila ---
    Dim idx2Keys() As String
    Dim idx2Rows() As Long
    Dim nIdx2      As Long
    nIdx2 = lastRow2 - 1
    ReDim idx2Keys(1 To nIdx2)
    ReDim idx2Rows(1 To nIdx2)
    Dim e As Long
    For e = 1 To nIdx2
        idx2Keys(e) = CStr(ws2.Cells(e + 1, colEmp2).Value)
        idx2Rows(e) = e + 1
    Next e

    ' --- Union ordenada de todos los Employee IDs ---
    Dim allIDs() As String
    Dim nAll     As Long
    Dim fila     As Long
    Dim found    As Boolean
    Dim k        As Long
    nAll = 0

    For fila = 2 To lastRow1
        Dim id1 As String
        id1 = CStr(ws1.Cells(fila, colEmp1).Value)
        If id1 <> "" Then
            found = False
            For k = 1 To nAll
                If allIDs(k) = id1 Then found = True: Exit For
            Next k
            If Not found Then
                nAll = nAll + 1
                ReDim Preserve allIDs(1 To nAll)
                allIDs(nAll) = id1
            End If
        End If
    Next fila

    For e = 1 To nIdx2
        Dim id2 As String
        id2 = idx2Keys(e)
        If id2 <> "" Then
            found = False
            For k = 1 To nAll
                If allIDs(k) = id2 Then found = True: Exit For
            Next k
            If Not found Then
                nAll = nAll + 1
                ReDim Preserve allIDs(1 To nAll)
                allIDs(nAll) = id2
            End If
        End If
    Next e

    ' Ordenar allIDs (burbuja)
    Dim tmp As String
    Dim j   As Long
    For k = 1 To nAll - 1
        For j = 1 To nAll - k
            If allIDs(j) > allIDs(j + 1) Then
                tmp = allIDs(j)
                allIDs(j) = allIDs(j + 1)
                allIDs(j + 1) = tmp
            End If
        Next j
    Next k

    ' --- Borrar COMPARACION anterior ---
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("COMPARACION").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' --- Crear hoja COMPARACION ---
    Dim wsC As Worksheet
    Set wsC = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsC.Name = "COMPARACION"

    ' --- Cabeceras ---
    Dim colC     As Long
    Dim cabecera As String

    For col = 1 To maxCol
        colC = (col - 1) * 2 + 1
        If col <= lastCol1 Then
            cabecera = CStr(ws1.Cells(1, col).Value)
        ElseIf col <= lastCol2 Then
            cabecera = CStr(ws2.Cells(1, col).Value)
        Else
            cabecera = "Campo" & col
        End If
        wsC.Range(wsC.Cells(1, colC), wsC.Cells(1, colC + 1)).Merge
        wsC.Cells(1, colC).Value = cabecera
        wsC.Cells(2, colC).Value = "v1"
        wsC.Cells(2, colC + 1).Value = "v2"
    Next col

    Dim colDif As Long
    colDif = maxCol * 2 + 1
    wsC.Cells(1, colDif).Value = "DIFERENTE"
    wsC.Cells(2, colDif).Value = ""

    ' Formato fila 1
    With wsC.Rows(1)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 20
    End With
    ' Formato fila 2
    With wsC.Rows(2)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(41, 128, 185)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 18
    End With
    wsC.Cells(2, colDif).Interior.Color = RGB(31, 78, 121)

    ' --- Escribir datos cruzados por Employee ID ---
    Application.ScreenUpdating = False

    Dim filaC   As Long
    Dim r1      As Long
    Dim r2      As Long
    Dim v1      As String
    Dim v2      As String
    Dim difFila As Boolean
    Dim estado  As String
    filaC = 3

    For k = 1 To nAll
        Dim empID As String
        empID = allIDs(k)

        ' Buscar en ws1
        r1 = 0
        For fila = 2 To lastRow1
            If CStr(ws1.Cells(fila, colEmp1).Value) = empID Then
                r1 = fila
                Exit For
            End If
        Next fila

        ' Buscar en ws2
        r2 = 0
        For e = 1 To nIdx2
            If idx2Keys(e) = empID Then
                r2 = idx2Rows(e)
                Exit For
            End If
        Next e

        ' Estado inicial
        difFila = False
        If r1 = 0 Then
            estado = "SOLO EN V2"
            difFila = True
        ElseIf r2 = 0 Then
            estado = "SOLO EN V1"
            difFila = True
        Else
            estado = "NO"
        End If

        ' Escribir valores
        For col = 1 To maxCol
            colC = (col - 1) * 2 + 1
            v1 = ""
            v2 = ""
            If r1 > 0 And col <= lastCol1 Then v1 = CStr(ws1.Cells(r1, col).Value)
            If r2 > 0 And col <= lastCol2 Then v2 = CStr(ws2.Cells(r2, col).Value)
            wsC.Cells(filaC, colC).Value = v1
            wsC.Cells(filaC, colC + 1).Value = v2
            If estado = "NO" And v1 <> v2 Then
                estado = "SI"
                difFila = True
            End If
        Next col

        ' Marcar fila segun estado
        If estado = "SI" Then
            ' Fondo blanco, celdas v2 distintas en rojo oscuro
            With wsC.Cells(filaC, colDif)
                .Value = "SI"
                .Font.Bold = True
                .Font.Color = RGB(192, 57, 43)
            End With
            For col = 1 To maxCol
                colC = (col - 1) * 2 + 1
                If CStr(wsC.Cells(filaC, colC).Value) <> _
                   CStr(wsC.Cells(filaC, colC + 1).Value) Then
                    With wsC.Cells(filaC, colC + 1)
                        .Interior.Color = RGB(139, 0, 0)
                        .Font.Color = RGB(255, 255, 255)
                        .Font.Bold = True
                    End With
                End If
            Next col

        ElseIf estado = "SOLO EN V1" Or estado = "SOLO EN V2" Then
            ' Fondo azul suave + tachado en datos, columna DIFERENTE vacia
            wsC.Rows(filaC).Interior.Color = RGB(213, 229, 242)
            With wsC.Range(wsC.Cells(filaC, 1), wsC.Cells(filaC, colDif - 1)).Font
                .Strikethrough = True
                .Color = RGB(80, 80, 80)
            End With
            With wsC.Cells(filaC, colDif)
                .Value = ""
                .Font.Strikethrough = False
            End With

        Else
            ' Igual: fondo blanco, verde en DIFERENTE
            With wsC.Cells(filaC, colDif)
                .Value = "NO"
                .Font.Color = RGB(39, 174, 96)
            End With
        End If

        filaC = filaC + 1
    Next k

    ' --- Bordes gruesos al final de cada grupo (borde derecho de cada columna v2) ---
    Dim lastDataRow As Long
    lastDataRow = filaC - 1

    For col = 1 To maxCol
        colC = (col - 1) * 2 + 2   ' columna v2 de cada grupo
        With wsC.Range(wsC.Cells(1, colC), wsC.Cells(lastDataRow, colC)).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(31, 78, 121)
        End With
    Next col

    With wsC.Range(wsC.Cells(1, colDif), wsC.Cells(lastDataRow, colDif)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(31, 78, 121)
    End With

    ' --- Autoajuste de ancho (minimo 8, maximo 40) ---
    wsC.Cells.EntireColumn.AutoFit
    Dim cIdx As Long
    For cIdx = 1 To colDif
        With wsC.Columns(cIdx)
            If .ColumnWidth < 8 Then .ColumnWidth = 8
            If .ColumnWidth > 40 Then .ColumnWidth = 40
        End With
    Next cIdx

    ' --- Autofiltro en fila 2 (sin celdas fusionadas) ---
    wsC.Rows(2).AutoFilter

    ' --- Congelar las 2 filas de cabecera ---
    Application.ScreenUpdating = True
    wsC.Activate
    wsC.Rows(3).Select
    ActiveWindow.FreezePanes = True
    wsC.Range("A1").Select

    ' --- Resumen ---
    Dim totalSI     As Long
    Dim totalSoloV1 As Long
    Dim totalSoloV2 As Long
    totalSI     = Application.WorksheetFunction.CountIf(wsC.Columns(colDif), "SI")
    totalSoloV1 = Application.WorksheetFunction.CountIf(wsC.Columns(colDif), "SOLO EN V1")
    totalSoloV2 = Application.WorksheetFunction.CountIf(wsC.Columns(colDif), "SOLO EN V2")

    MsgBox "Comparacion completada." & vbCrLf & vbCrLf & _
           "  Registros totales : " & nAll & vbCrLf & _
           "  Campos diferentes : " & totalSI & vbCrLf & _
           "  Solo en v1        : " & totalSoloV1 & vbCrLf & _
           "  Solo en v2        : " & totalSoloV2 & vbCrLf & _
           "  Iguales           : " & (nAll - totalSI - totalSoloV1 - totalSoloV2) & vbCrLf & vbCrLf & _
           "Filtra columna DIFERENTE para ver los cambios.", _
           vbInformation, "Resultado"
End Sub
