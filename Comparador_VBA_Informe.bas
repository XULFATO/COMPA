' ============================================================
' COMPARADOR DE EXCELS
' Cruza dos ficheros por * Employee ID y saca diferencias.
' Cada campo sale doble: BULK ANTERIOR (antes) y BULK ACTUAL (ahora).
' Botones: IMPORTAR HOY 1/2 - COMPARAR - INFORME - BORRAR TODO
' ============================================================


Sub ImportarHoy1()
    ImportarHoja 1
End Sub

Sub ImportarHoy2()
    ImportarHoja 2
End Sub


' ============================================================
' BORRAR TODO
' Limpia todo y deja solo el MENU, para empezar de cero.
' ============================================================
Sub BorrarTodo()

    ' Preguntamos antes, no vaya a ser que se pulse sin querer
    Dim resp As Integer
    resp = MsgBox("Se eliminaran todas las hojas excepto MENU." & vbCrLf & _
                  "Esto incluye las importadas y la comparacion." & vbCrLf & vbCrLf & _
                  "Continuar?", vbQuestion + vbYesNo, "Confirmar borrado")
    If resp = vbNo Then Exit Sub

    Dim wsMenu As Worksheet
    Set wsMenu = ThisWorkbook.Worksheets("MENU")

    ' Borramos los nombres guardados en J1 y J2
    wsMenu.Range("J1").Value = ""
    wsMenu.Range("J2").Value = ""

    ' Eliminamos todo lo que no sea MENU.
    ' Recorremos en orden inverso para no saltarnos hojas al borrar.
    Application.DisplayAlerts = False
    Dim iWs As Integer
    For iWs = ThisWorkbook.Worksheets.Count To 1 Step -1
        If ThisWorkbook.Worksheets(iWs).Name <> "MENU" Then
            ThisWorkbook.Worksheets(iWs).Delete
        End If
    Next iWs
    Application.DisplayAlerts = True

    wsMenu.Activate
    MsgBox "Listo. Solo queda la hoja MENU.", vbInformation, "Borrado completado"
End Sub


' ============================================================
' IMPORTAR HOJA (slot 1 o 2)
' Lista los ficheros abiertos, el usuario elige cual,
' se copia la primera hoja con el nombre + sufijo v1 o v2.
' ============================================================
Sub ImportarHoja(slot As Integer)

    ' Referencia al MENU al inicio del todo, antes de cualquier Copy
    ' porque despues Excel puede cambiar el libro activo y liarnos
    Dim wsMenu As Worksheet
    Set wsMenu = ThisWorkbook.Worksheets("MENU")

    ' Miramos que ficheros hay abiertos ademas de este
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

    ' Mostramos la lista y pedimos que elijan. Si cancelan salimos.
    Dim lista As String
    lista = "Ficheros Excel abiertos:" & vbCrLf & vbCrLf
    For i = 0 To n - 1
        lista = lista & "  " & (i + 1) & "  ->  " & nombres(i) & vbCrLf
    Next i
    lista = lista & vbCrLf & "Escribe el numero del fichero:"

    Dim respWB As Variant
    respWB = Application.InputBox(lista, "Importar HOY " & slot, Type:=1)

    ' Cuando se cancela un InputBox numerico devuelve False (booleano)
    If VarType(respWB) = vbBoolean Then Exit Sub

    Dim idxWB As Integer
    idxWB = CInt(respWB) - 1
    If idxWB < 0 Or idxWB >= n Then
        MsgBox "Numero fuera de rango (1 a " & n & ").", vbExclamation
        Exit Sub
    End If

    Dim wbOrigen As Workbook
    Set wbOrigen = Application.Workbooks(nombres(idxWB))

    ' Cogemos siempre la primera hoja, que es la PAGE 1 que nos interesa
    Dim wsOrigen As Worksheet
    Set wsOrigen = wbOrigen.Worksheets(1)

    ' Nombre de la pestana = nombre original + " v1" o " v2"
    ' Recortamos si es muy largo y quitamos caracteres ilegales
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

    ' Si ya existia esa pestana la borramos antes de copiar
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(nomHoja).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Guardamos referencia al libro ANTES del Copy
    ' (el Copy puede cambiar el libro activo y perder ThisWorkbook)
    Dim wbThis     As Workbook
    Dim totalHojas As Integer
    Set wbThis = ThisWorkbook
    totalHojas = wbThis.Worksheets.Count
    wsOrigen.Copy After:=wbThis.Worksheets(totalHojas)
    wbThis.Worksheets(wbThis.Worksheets.Count).Name = nomHoja

    ' Guardamos el nombre en J1 o J2 del MENU para que CompararHojas lo encuentre
    If slot = 1 Then
        wsMenu.Range("J1").Value = nomHoja
    Else
        wsMenu.Range("J2").Value = nomHoja
    End If

    wsMenu.Activate
End Sub


' ============================================================
' COMPARAR HOJAS
' Lee J1/J2, cruza por * Employee ID, genera COMPARACION.
' FILTRO (col A): IGUALES / DIFERENTES / SOLO EN V1 / SOLO EN V2
' ============================================================
Sub CompararHojas()

    ' MENU al inicio, antes de todo
    Dim wsMenu As Worksheet
    Set wsMenu = ThisWorkbook.Worksheets("MENU")

    ' Nombres de las hojas que se importaron (los guarda ImportarHoja)
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

    ' Hasta donde llegan los datos en cada hoja
    Dim lastRow1 As Long, lastCol1 As Long
    Dim lastRow2 As Long, lastCol2 As Long
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastCol1 = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    lastCol2 = ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column
    Dim maxCol As Long
    maxCol = Application.Max(lastCol1, lastCol2)

    ' Buscamos * Employee ID en cada hoja (comparamos en minusculas)
    Dim colEmp1 As Long
    Dim colEmp2 As Long
    Dim col     As Long
    colEmp1 = 0
    colEmp2 = 0

    For col = 1 To lastCol1
        If LCase(Trim(CStr(ws1.Cells(1, col).Value))) = "* employee id" Then
            colEmp1 = col
            Exit For
        End If
    Next col
    For col = 1 To lastCol2
        If LCase(Trim(CStr(ws2.Cells(1, col).Value))) = "* employee id" Then
            colEmp2 = col
            Exit For
        End If
    Next col

    If colEmp1 = 0 Or colEmp2 = 0 Then
        MsgBox "No se encuentra '* Employee ID' en " & _
               IIf(colEmp1 = 0, nomH1, nomH2) & "." & vbCrLf & _
               "Verifica que el nombre es exactamente '* Employee ID'.", _
               vbExclamation, "Columna no encontrada"
        Exit Sub
    End If

    ' Indice de ws2: dos arrays paralelos (IDs y filas) para busqueda rapida
    Dim idx2Keys() As String
    Dim idx2Rows() As Long
    Dim nIdx2      As Long
    nIdx2 = lastRow2 - 1
    If nIdx2 < 1 Then
        MsgBox "La hoja HOY 2 no tiene datos (solo cabecera).", vbExclamation
        Exit Sub
    End If
    ReDim idx2Keys(1 To nIdx2)
    ReDim idx2Rows(1 To nIdx2)
    Dim e As Long
    For e = 1 To nIdx2
        idx2Keys(e) = CStr(ws2.Cells(e + 1, colEmp2).Value)
        idx2Rows(e) = e + 1
    Next e

    ' Union de IDs de los dos ficheros sin duplicados
    Dim allIDs() As String
    Dim nAll     As Long
    Dim fila     As Long
    Dim found    As Boolean
    Dim k        As Long
    nAll = 0

    Dim id1 As String
    ' IDs de ws1
    For fila = 2 To lastRow1
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

    Dim id2 As String
    ' IDs de ws2 que no esten ya
    For e = 1 To nIdx2
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

    ' Ordenar IDs (burbuja, suficiente para este volumen)
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

    ' Borramos COMPARACION anterior si existe
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("COMPARACION").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Dim wsC As Worksheet
    Set wsC = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsC.Name = "COMPARACION"
    Application.ScreenUpdating = False   ' apagamos ya, antes de escribir cabeceras

    ' ----------------------------------------------------------
    ' Cabeceras: FILTRO en col 1, luego BULK ANTERIOR/BULK ACTUAL por cada campo
    ' colC = (col-1)*2 + 2 -> BULK ANTERIOR del campo col
    ' ----------------------------------------------------------

    ' FILTRO fusionado en filas 1 y 2 de la columna A
    wsC.Range(wsC.Cells(1, 1), wsC.Cells(2, 1)).Merge
    wsC.Cells(1, 1).Value = "FILTRO"

    Dim colC     As Long
    Dim cabecera As String
    For col = 1 To maxCol
        colC = (col - 1) * 2 + 2

        If col <= lastCol1 Then
            cabecera = CStr(ws1.Cells(1, col).Value)
        ElseIf col <= lastCol2 Then
            cabecera = CStr(ws2.Cells(1, col).Value)
        Else
            cabecera = "Campo" & col
        End If

        wsC.Range(wsC.Cells(1, colC), wsC.Cells(1, colC + 1)).Merge
        wsC.Cells(1, colC).Value = cabecera
        wsC.Cells(2, colC).Value = "BULK ANTERIOR"
        wsC.Cells(2, colC + 1).Value = "BULK ACTUAL"
    Next col

    ' Fila 1: azul oscuro en todo, naranja en FILTRO
    With wsC.Rows(1)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 20
    End With

    ' Fila 2: azul medio en todo, naranja en FILTRO
    With wsC.Rows(2)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(41, 128, 185)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 18
    End With

    ' La columna FILTRO en naranja para que destaque del resto
    wsC.Cells(1, 1).Interior.Color = RGB(197, 90, 17)
    wsC.Cells(2, 1).Interior.Color = RGB(197, 90, 17)

    ' ----------------------------------------------------------
    ' Bucle principal: un empleado por iteracion
    ' ----------------------------------------------------------

    Dim filaC      As Long
    Dim r1         As Long
    Dim r2         As Long
    Dim v1         As String
    Dim v2         As String
    Dim difFila    As Boolean
    Dim estado     As String
    Dim cntSI      As Long
    Dim cntSoloV1  As Long
    Dim cntSoloV2  As Long
    Dim cntIguales As Long
    Dim lastDataCol As Long
    lastDataCol = maxCol * 2 + 1   ' ultima col BULK ACTUAL del ultimo campo
    filaC = 3
    cntSI = 0: cntSoloV1 = 0: cntSoloV2 = 0: cntIguales = 0

    Dim empID As String
    For k = 1 To nAll
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

        ' Estado inicial segun donde aparece
        difFila = False
        If r1 = 0 Then
            ' Alta nueva en ws2
            estado = "SOLO EN V2"
            difFila = True
        ElseIf r2 = 0 Then
            ' Estaba en ws1, ya no esta en ws2
            estado = "SOLO EN V1"
            difFila = True
        Else
            ' En los dos, asumimos iguales hasta que encontremos algo distinto
            estado = "NO"
        End If

        ' Escribir valores BULK ANTERIOR/ACTUAL (col 1 es FILTRO, datos desde col 2)
        For col = 1 To maxCol
            colC = (col - 1) * 2 + 2
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

        ' Marcar segun estado
        If estado = "SI" Then
            ' Fondo blanco en datos. En FILTRO rojo oscuro.
            ' BULK ACTUAL que cambio: rojo oscuro. BULK ANTERIOR su pareja: rojo clarito.
            With wsC.Cells(filaC, 1)
                .Value = "DIFERENTES"
                .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(192, 57, 43)
            End With
            For col = 1 To maxCol
                colC = (col - 1) * 2 + 2
                If CStr(wsC.Cells(filaC, colC).Value) <> _
                   CStr(wsC.Cells(filaC, colC + 1).Value) Then
                    ' BULK ACTUAL (el nuevo valor): rojo oscuro
                    With wsC.Cells(filaC, colC + 1)
                        .Interior.Color = RGB(139, 0, 0)
                        .Font.Color = RGB(255, 255, 255)
                        .Font.Bold = True
                    End With
                    ' BULK ANTERIOR (el valor antiguo): rojo clarito, pa ver la pareja
                    With wsC.Cells(filaC, colC)
                        .Interior.Color = RGB(255, 199, 199)
                        .Font.Bold = True
                    End With
                End If
            Next col

        ElseIf estado = "SOLO EN V1" Then
            ' Ya no esta. Azul suave + tachado en los datos.
            wsC.Range(wsC.Cells(filaC, 2), wsC.Cells(filaC, lastDataCol)).Interior.Color = RGB(213, 229, 242)
            With wsC.Range(wsC.Cells(filaC, 2), wsC.Cells(filaC, lastDataCol)).Font
                .Strikethrough = True
                .Color = RGB(80, 80, 80)
            End With
            With wsC.Cells(filaC, 1)
                .Value = "SOLO EN V1"
                .Font.Bold = True
                .Font.Strikethrough = False
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(31, 78, 121)
            End With

        ElseIf estado = "SOLO EN V2" Then
            ' Alta nueva. Verde suave en datos, verde oscuro en FILTRO.
            wsC.Range(wsC.Cells(filaC, 2), wsC.Cells(filaC, lastDataCol)).Interior.Color = RGB(198, 239, 206)
            With wsC.Cells(filaC, 1)
                .Value = "SOLO EN V2"
                .Font.Bold = True
                .Font.Strikethrough = False
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(21, 101, 71)
            End With

        Else
            ' Todo igual. Verde suave en datos, verde oscuro en FILTRO.
            wsC.Range(wsC.Cells(filaC, 2), wsC.Cells(filaC, lastDataCol)).Interior.Color = RGB(198, 239, 206)
            With wsC.Cells(filaC, 1)
                .Value = "IGUALES"
                .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(39, 120, 56)
            End With
        End If

        ' Contadores
        If estado = "SI" Then
            cntSI = cntSI + 1
        ElseIf estado = "SOLO EN V1" Then
            cntSoloV1 = cntSoloV1 + 1
        ElseIf estado = "SOLO EN V2" Then
            cntSoloV2 = cntSoloV2 + 1
        Else
            cntIguales = cntIguales + 1
        End If

        filaC = filaC + 1
    Next k

    ' Borde derecho de columna FILTRO
    Dim lastDataRow As Long
    lastDataRow = filaC - 1
    With wsC.Range(wsC.Cells(1, 1), wsC.Cells(lastDataRow, 1)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(31, 78, 121)
    End With

    ' Borde derecho de cada BULK ACTUAL pa separar grupos
    For col = 1 To maxCol
        colC = (col - 1) * 2 + 3
        With wsC.Range(wsC.Cells(1, colC), wsC.Cells(lastDataRow, colC)).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(31, 78, 121)
        End With
    Next col

    ' Autoajuste de ancho (min 8, max 40)
    wsC.Cells.EntireColumn.AutoFit
    Dim cIdx As Long
    For cIdx = 1 To lastDataCol
        With wsC.Columns(cIdx)
            If .ColumnWidth < 8  Then .ColumnWidth = 8
            If .ColumnWidth > 40 Then .ColumnWidth = 40
        End With
    Next cIdx

    ' AutoFilter desde A1
    wsC.Range("A1").AutoFilter

    ' Congelar cabeceras (filas 1-2) y columna FILTRO (col A)
    ' Seleccionamos B3: todo lo de arriba y a la izquierda queda bloqueado
    Application.ScreenUpdating = True
    wsC.Activate
    wsC.Range("B3").Select
    ActiveWindow.FreezePanes = True
    wsC.Range("A1").Select

    ' Resumen final
    MsgBox "Comparacion completada." & vbCrLf & vbCrLf & _
           "  Registros totales : " & nAll & vbCrLf & _
           "  DIFERENTES        : " & cntSI & vbCrLf & _
           "  IGUALES           : " & cntIguales & vbCrLf & _
           "  SOLO EN V1        : " & cntSoloV1 & vbCrLf & _
           "  SOLO EN V2        : " & cntSoloV2 & vbCrLf & vbCrLf & _
           "Filtra columna FILTRO para ver el resultado por fila.", _
           vbInformation, "Resultado"

End Sub


' ============================================================
' INFORME
' Crea 4 pestañas a partir de COMPARACION, una por tipo.
' Mismo layout y colores. Requiere haber comparado antes.
' ============================================================
Sub Informe()

    Dim wsC As Worksheet
    On Error Resume Next
    Set wsC = ThisWorkbook.Worksheets("COMPARACION")
    On Error GoTo 0

    If wsC Is Nothing Then
        MsgBox "Primero ejecuta COMPARAR.", vbExclamation, "Falta COMPARACION"
        Exit Sub
    End If

    Dim lastRow As Long, lastCol As Long
    lastRow = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
    lastCol = wsC.Cells(2, wsC.Columns.Count).End(xlToLeft).Column

    If lastRow < 3 Then
        MsgBox "La hoja COMPARACION esta vacia.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim tabNames(3)   As String
    Dim tabFiltros(3) As String
    tabNames(0) = "DIFERENCIAS": tabFiltros(0) = "DIFERENTES"
    tabNames(1) = "IGUALES":     tabFiltros(1) = "IGUALES"
    tabNames(2) = "SOLO EN V1":  tabFiltros(2) = "SOLO EN V1"
    tabNames(3) = "SOLO EN V2":  tabFiltros(3) = "SOLO EN V2"

    ' Todas las Dim fuera del bucle para evitar conflictos con el flujo
    Dim wsT         As Worksheet
    Dim t           As Integer
    Dim filaOrigen  As Long
    Dim filaDestino As Long
    Dim cntTipo     As Long
    Dim ci          As Long
    Dim resumen     As String
    resumen = ""

    For t = 0 To 3

        ' Contar cuantas filas hay de este tipo ANTES de crear la pestana
        cntTipo = 0
        For filaOrigen = 3 To lastRow
            If CStr(wsC.Cells(filaOrigen, 1).Value) = tabFiltros(t) Then
                cntTipo = cntTipo + 1
            End If
        Next filaOrigen

        ' Si no hay nada de este tipo nos lo saltamos
        If cntTipo = 0 Then GoTo SiguientePestana

        ' Borrar si ya existe de antes
        On Error Resume Next
        ThisWorkbook.Worksheets(tabNames(t)).Delete
        On Error GoTo 0

        ' Crear al final
        Set wsT = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsT.Name = tabNames(t)

        ' Copiar cabeceras exactas de COMPARACION
        wsC.Rows(1).Copy wsT.Rows(1)
        wsC.Rows(2).Copy wsT.Rows(2)

        ' Copiar solo las filas de este tipo
        filaDestino = 3
        For filaOrigen = 3 To lastRow
            If CStr(wsC.Cells(filaOrigen, 1).Value) = tabFiltros(t) Then
                wsC.Rows(filaOrigen).Copy wsT.Rows(filaDestino)
                filaDestino = filaDestino + 1
            End If
        Next filaOrigen

        ' Autofit con limites
        wsT.Cells.EntireColumn.AutoFit
        For ci = 1 To lastCol
            With wsT.Columns(ci)
                If .ColumnWidth < 8  Then .ColumnWidth = 8
                If .ColumnWidth > 40 Then .ColumnWidth = 40
            End With
        Next ci

        ' AutoFilter en fila 2
        wsT.Rows(2).AutoFilter

        ' Congelar filas 1-2 y columna FILTRO
        wsT.Activate
        wsT.Range("B3").Select
        ActiveWindow.FreezePanes = True
        wsT.Range("A1").Select

        ' Color de pestana segun tipo
        Select Case t
            Case 0: wsT.Tab.Color = RGB(192, 57, 43)
            Case 1: wsT.Tab.Color = RGB(39, 120, 56)
            Case 2: wsT.Tab.Color = RGB(31, 78, 121)
            Case 3: wsT.Tab.Color = RGB(21, 101, 71)
        End Select

        ' Acumular en el resumen solo las que se han creado
        resumen = resumen & "  - " & tabNames(t) & " (" & cntTipo & " filas)" & vbCrLf

SiguientePestana:
    Next t

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    wsC.Activate
    wsC.Range("A1").Select

    If resumen = "" Then
        MsgBox "No se ha generado ninguna pestana (COMPARACION vacia).", _
               vbInformation, "Sin datos"
    Else
        MsgBox "Pestanas generadas:" & vbCrLf & vbCrLf & resumen, _
               vbInformation, "Listo"
    End If
End Sub
