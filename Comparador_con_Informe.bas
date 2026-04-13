' ================================================================
' COMPARADOR DE EXCELS - con Informe
' ================================================================
' Para que funcione necesitas:
'   1. Dos ficheros Excel abiertos con la misma estructura
'   2. Una columna llamada exactamente "* Employee ID" en cada uno
'   3. Ejecutar los subs en este orden: Importar 1 -> Importar 2
'      -> Comparar -> (opcional) Informe -> (opcional) Exportar
'
' La clave de cruce es * Employee ID, no la posicion de la fila.
' Cada campo original aparece doble: BULK ANTERIOR y BULK ACTUAL.
' La columna FILTRO (col A) indica el estado de cada registro:
'   IGUALES    -> sin cambios
'   DIFERENTES -> hay campos que han cambiado
'   SOLO EN V1 -> estaba en el fichero antiguo, ya no esta
'   SOLO EN V2 -> alta nueva, no estaba antes
'
' Los nombres de las hojas importadas se guardan en celdas
' ocultas J1 y J2 de la hoja MENU para que los subs se
' comuniquen entre si sin necesidad de variables globales.
' ================================================================


' ----------------------------------------------------------------
' Wrappers de boton - llaman a ImportarHoja con el slot correcto
' ----------------------------------------------------------------
Sub ImportarHoy1()
    ImportarHoja 1
End Sub

Sub ImportarHoy2()
    ImportarHoja 2
End Sub


' ================================================================
' BORRAR TODO
' Elimina todas las hojas menos MENU y limpia J1/J2.
' Util para empezar de cero sin cerrar el fichero.
' Pide confirmacion antes de borrar, no vaya a ser que
' alguien le de sin querer.
' ================================================================
Sub BorrarTodo()

    Dim resp As Integer
    resp = MsgBox("Se eliminaran todas las hojas excepto MENU." & vbCrLf & _
                  "Esto incluye las importadas y la comparacion." & vbCrLf & vbCrLf & _
                  "Continuar?", vbQuestion + vbYesNo, "Confirmar borrado")
    If resp = vbNo Then Exit Sub

    ' Cogemos MENU antes de ponernos a borrar cosas
    Dim wsMenu As Worksheet
    Set wsMenu = ThisWorkbook.Worksheets("MENU")

    ' Limpiamos las referencias a las hojas importadas
    wsMenu.Range("J1").Value = ""
    wsMenu.Range("J2").Value = ""

    ' Recorremos en orden inverso: si borramos hacia adelante
    ' el indice se descuadra y nos saltamos hojas
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


' ================================================================
' IMPORTAR HOJA
' Parametro slot: 1 para el fichero antiguo, 2 para el actual.
'
' Busca los ficheros Excel abiertos (menos este), muestra
' la lista, el usuario elige uno, y se copia su primera hoja
' aqui dentro con el nombre original + sufijo " v1" o " v2".
'
' El nombre de la hoja copiada se guarda en MENU!J1 o J2
' para que CompararHojas sepa donde buscar despues.
'
' Por que guardamos wbThis antes del Copy:
'   wsOrigen.Copy abre temporalmente el libro copiado como
'   activo, y en ese momento ThisWorkbook puede apuntar al
'   sitio equivocado. Guardando la referencia antes evitamos
'   el problema.
' ================================================================
Sub ImportarHoja(slot As Integer)

    ' Referencia al MENU al inicio, antes de cualquier Copy
    Dim wsMenu As Worksheet
    Set wsMenu = ThisWorkbook.Worksheets("MENU")

    ' --- Paso 1: listar ficheros abiertos (excluir este) ---
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

    ' --- Paso 2: pedir al usuario que elija fichero ---
    ' Application.InputBox con Type:=1 fuerza numero entero
    ' y devuelve False (booleano) si se pulsa Cancelar
    Dim lista As String
    lista = "Ficheros Excel abiertos:" & vbCrLf & vbCrLf
    For i = 0 To n - 1
        lista = lista & "  " & (i + 1) & "  ->  " & nombres(i) & vbCrLf
    Next i
    lista = lista & vbCrLf & "Escribe el numero del fichero:"

    Dim respWB As Variant
    respWB = Application.InputBox(lista, "Importar HOY " & slot, Type:=1)

    If VarType(respWB) = vbBoolean Then Exit Sub   ' Cancelar

    Dim idxWB As Integer
    idxWB = CInt(respWB) - 1
    If idxWB < 0 Or idxWB >= n Then
        MsgBox "Numero fuera de rango (1 a " & n & ").", vbExclamation
        Exit Sub
    End If

    Dim wbOrigen As Workbook
    Set wbOrigen = Application.Workbooks(nombres(idxWB))

    ' --- Paso 3: coger siempre la primera hoja (PAGE 1) ---
    Dim wsOrigen As Worksheet
    Set wsOrigen = wbOrigen.Worksheets(1)

    ' --- Paso 4: construir nombre de la pestana destino ---
    ' Formato: nombre_hoja_origen + " v1" o " v2"
    ' Recortamos a 25 chars y eliminamos caracteres ilegales
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

    ' --- Paso 5: borrar la importacion anterior del mismo slot ---
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(nomHoja).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' --- Paso 6: copiar la hoja ---
    ' Guardamos wbThis y el count ANTES del Copy para no perder
    ' la referencia a ThisWorkbook cuando Excel cambie el libro activo
    Dim wbThis     As Workbook
    Dim totalHojas As Integer
    Set wbThis = ThisWorkbook
    totalHojas = wbThis.Worksheets.Count

    wsOrigen.Copy After:=wbThis.Worksheets(totalHojas)
    wbThis.Worksheets(wbThis.Worksheets.Count).Name = nomHoja

    ' --- Paso 7: guardar el nombre en MENU para que CompararHojas lo encuentre ---
    If slot = 1 Then
        wsMenu.Range("J1").Value = nomHoja
    Else
        wsMenu.Range("J2").Value = nomHoja
    End If

    ' Volvemos al MENU sin mensaje, mas limpio
    wsMenu.Activate
End Sub


' ================================================================
' COMPARAR HOJAS
'
' Lee los nombres de las hojas desde MENU J1 y J2.
' Cruza los registros por * Employee ID (no por posicion de fila).
' Genera la hoja COMPARACION con:
'   - Col A: FILTRO con el estado de cada registro
'   - Por cada campo: columna BULK ANTERIOR + columna BULK ACTUAL
'
' Logica de cruce:
'   Si el ID esta en los dos ficheros -> comparamos campo a campo
'   Si solo esta en V1 -> SOLO EN V1 (baja)
'   Si solo esta en V2 -> SOLO EN V2 (alta)
'
' Colores en COMPARACION:
'   FILTRO naranja para que destaque de las cabeceras azules
'   DIFERENTES: celda BULK ACTUAL en rojo oscuro, ANTERIOR en rojo clarito
'   SOLO EN V1: fondo azul suave + texto tachado
'   SOLO EN V2: fondo verde suave
'   IGUALES: fondo verde suave
' ================================================================
Sub CompararHojas()

    On Error GoTo CompararError

    ' MENU al inicio del todo
    Dim wsMenu As Worksheet
    Set wsMenu = ThisWorkbook.Worksheets("MENU")

    ' Leemos los nombres de las hojas importadas (guardados por ImportarHoja)
    Dim nomH1 As String
    Dim nomH2 As String
    nomH1 = Trim(wsMenu.Range("J1").Value)
    nomH2 = Trim(wsMenu.Range("J2").Value)

    If nomH1 = "" Or nomH2 = "" Then
        MsgBox "Importa primero las dos hojas (HOY 1 y HOY 2).", _
               vbExclamation, "Faltan hojas"
        Exit Sub
    End If

    ' Buscamos las hojas dentro de este libro
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets(nomH1)
    Set ws2 = ThisWorkbook.Worksheets(nomH2)
    On Error GoTo CompararError

    If ws1 Is Nothing Then
        MsgBox "No encuentro la hoja HOY 1:" & vbCrLf & nomH1, vbCritical
        Exit Sub
    End If
    If ws2 Is Nothing Then
        MsgBox "No encuentro la hoja HOY 2:" & vbCrLf & nomH2, vbCritical
        Exit Sub
    End If

    ' Calculamos hasta donde llegan los datos en cada hoja
    Dim lastRow1 As Long, lastCol1 As Long
    Dim lastRow2 As Long, lastCol2 As Long
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastCol1 = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    lastCol2 = ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column

    ' El maximo de columnas es el mayor de los dos, por si tienen columnas distintas
    Dim maxCol As Long
    maxCol = Application.Max(lastCol1, lastCol2)

    ' Buscamos la columna * Employee ID en cada hoja (insensible a mayusculas)
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

    ' Construimos un indice de ws2 para busqueda rapida por ID:
    ' dos arrays paralelos idx2Keys (IDs) e idx2Rows (filas reales en ws2)
    Dim idx2Keys() As String
    Dim idx2Rows() As Long
    Dim nIdx2      As Long
    nIdx2 = lastRow2 - 1   ' filas de datos, sin la cabecera
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

    ' Construimos la lista completa de IDs: union de ws1 y ws2 sin duplicados
    ' Primero los de ws1, luego los de ws2 que no esten ya
    Dim allIDs() As String
    Dim nAll     As Long
    Dim fila     As Long
    Dim found    As Boolean
    Dim k        As Long
    nAll = 0

    Dim id1 As String
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

    ' Ordenamos los IDs con burbuja para que el resultado salga ordenado
    ' Para el volumen de empleados que manejamos esto es mas que suficiente
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

    ' Borramos COMPARACION anterior si existia de una ejecucion previa
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("COMPARACION").Delete
    On Error GoTo CompararError
    Application.DisplayAlerts = True

    ' Creamos la hoja COMPARACION al final del libro
    Dim wsC As Worksheet
    Set wsC = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsC.Name = "COMPARACION"

    ' Apagamos la pantalla antes de escribir cabeceras y datos
    Application.ScreenUpdating = False

    ' --- Construir cabeceras ---
    ' Fila 1: nombre del campo original, fusionado sobre BULK ANTERIOR y BULK ACTUAL
    ' Fila 2: etiquetas BULK ANTERIOR / BULK ACTUAL
    ' Col 1 siempre es FILTRO (fusionado en filas 1 y 2)
    ' Los campos empiezan en col 2: colC = (col-1)*2 + 2
    wsC.Range(wsC.Cells(1, 1), wsC.Cells(2, 1)).Merge
    wsC.Cells(1, 1).Value = "FILTRO"

    Dim colC     As Long
    Dim cabecera As String
    For col = 1 To maxCol
        colC = (col - 1) * 2 + 2

        ' Nombre del campo: de ws1 si existe ahi, si no de ws2
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

    ' Formato fila 1: azul oscuro, texto blanco
    With wsC.Rows(1)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 20
    End With

    ' Formato fila 2: azul medio, texto blanco
    With wsC.Rows(2)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(41, 128, 185)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 18
    End With

    ' Columna FILTRO en naranja para que destaque del resto de cabeceras azules
    wsC.Cells(1, 1).Interior.Color = RGB(197, 90, 17)
    wsC.Cells(2, 1).Interior.Color = RGB(197, 90, 17)

    ' --- Bucle principal: un ID de empleado por iteracion ---
    Dim filaC      As Long    ' fila actual en wsC
    Dim r1         As Long    ' fila del empleado en ws1 (0 si no esta)
    Dim r2         As Long    ' fila del empleado en ws2 (0 si no esta)
    Dim v1         As String  ' valor BULK ANTERIOR
    Dim v2         As String  ' valor BULK ACTUAL
    Dim difFila    As Boolean
    Dim estado     As String
    Dim cntSI      As Long
    Dim cntSoloV1  As Long
    Dim cntSoloV2  As Long
    Dim cntIguales As Long
    Dim lastDataCol As Long
    lastDataCol = maxCol * 2 + 1   ' ultima columna con datos (BULK ACTUAL del ultimo campo)
    filaC = 3   ' los datos empiezan en fila 3, filas 1-2 son cabeceras
    cntSI = 0: cntSoloV1 = 0: cntSoloV2 = 0: cntIguales = 0

    Dim empID As String
    For k = 1 To nAll
        empID = allIDs(k)

        ' Buscar la fila de este ID en ws1 (busqueda lineal)
        r1 = 0
        For fila = 2 To lastRow1
            If CStr(ws1.Cells(fila, colEmp1).Value) = empID Then
                r1 = fila
                Exit For
            End If
        Next fila

        ' Buscar en ws2 usando el indice que construimos antes
        r2 = 0
        For e = 1 To nIdx2
            If idx2Keys(e) = empID Then
                r2 = idx2Rows(e)
                Exit For
            End If
        Next e

        ' Determinar el estado inicial antes de comparar campos
        difFila = False
        If r1 = 0 Then
            estado = "SOLO EN V2"   ' no estaba en ws1, es una alta nueva
            difFila = True
        ElseIf r2 = 0 Then
            estado = "SOLO EN V1"   ' no esta en ws2, ha sido dado de baja
            difFila = True
        Else
            estado = "NO"   ' esta en los dos, asumimos iguales hasta comparar
        End If

        ' Escribir los valores de cada campo lado a lado
        ' Si el empleado no esta en uno de los ficheros esa columna queda vacia
        For col = 1 To maxCol
            colC = (col - 1) * 2 + 2
            v1 = ""
            v2 = ""
            If r1 > 0 And col <= lastCol1 Then v1 = CStr(ws1.Cells(r1, col).Value)
            If r2 > 0 And col <= lastCol2 Then v2 = CStr(ws2.Cells(r2, col).Value)
            wsC.Cells(filaC, colC).Value = v1
            wsC.Cells(filaC, colC + 1).Value = v2
            ' Si encontramos una diferencia cambiamos el estado
            ' Solo actualizamos si era "NO" para no machacar SOLO EN V1/V2
            If estado = "NO" And v1 <> v2 Then
                estado = "SI"
                difFila = True
            End If
        Next col

        ' --- Marcar la fila segun su estado ---
        If estado = "SI" Then
            ' Hay campos distintos entre los dos ficheros
            ' Fondo blanco en datos, FILTRO en rojo oscuro
            ' BULK ACTUAL que cambio: rojo oscuro / BULK ANTERIOR su pareja: rojo clarito
            With wsC.Cells(filaC, 1)
                .Value = "DIFERENTES"
                .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(192, 57, 43)
            End With
            ' Segundo recorrido para colorear celda a celda los cambios
            For col = 1 To maxCol
                colC = (col - 1) * 2 + 2
                If CStr(wsC.Cells(filaC, colC).Value) <> _
                   CStr(wsC.Cells(filaC, colC + 1).Value) Then
                    With wsC.Cells(filaC, colC + 1)   ' BULK ACTUAL: rojo oscuro
                        .Interior.Color = RGB(139, 0, 0)
                        .Font.Color = RGB(255, 255, 255)
                        .Font.Bold = True
                    End With
                    With wsC.Cells(filaC, colC)   ' BULK ANTERIOR: rojo clarito
                        .Interior.Color = RGB(255, 199, 199)
                        .Font.Bold = True
                    End With
                End If
            Next col

        ElseIf estado = "SOLO EN V1" Then
            ' Este empleado ya no existe en el fichero actual
            ' Fondo azul suave + tachado para que se vea que "ha desaparecido"
            wsC.Range(wsC.Cells(filaC, 2), wsC.Cells(filaC, lastDataCol)).Interior.Color = RGB(213, 229, 242)
            With wsC.Range(wsC.Cells(filaC, 2), wsC.Cells(filaC, lastDataCol)).Font
                .Strikethrough = True
                .Color = RGB(80, 80, 80)
            End With
            ' La celda FILTRO no la tachamos, tiene que leerse bien
            With wsC.Cells(filaC, 1)
                .Value = "SOLO EN V1"
                .Font.Bold = True
                .Font.Strikethrough = False
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(31, 78, 121)
            End With

        ElseIf estado = "SOLO EN V2" Then
            ' Alta nueva: no estaba en el fichero anterior
            ' Fondo verde suave, sin tachado
            wsC.Range(wsC.Cells(filaC, 2), wsC.Cells(filaC, lastDataCol)).Interior.Color = RGB(198, 239, 206)
            With wsC.Cells(filaC, 1)
                .Value = "SOLO EN V2"
                .Font.Bold = True
                .Font.Strikethrough = False
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(21, 101, 71)
            End With

        Else
            ' Todo igual entre los dos ficheros
            ' Fondo verde suave en datos, verde oscuro en FILTRO
            wsC.Range(wsC.Cells(filaC, 2), wsC.Cells(filaC, lastDataCol)).Interior.Color = RGB(198, 239, 206)
            With wsC.Cells(filaC, 1)
                .Value = "IGUALES"
                .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(39, 120, 56)
            End With
        End If

        ' Acumulamos en los contadores para el resumen final
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

    ' --- Formato final de la hoja COMPARACION ---

    ' Borde grueso a la derecha de la columna FILTRO para separarla visualmente
    Dim lastDataRow As Long
    lastDataRow = filaC - 1
    With wsC.Range(wsC.Cells(1, 1), wsC.Cells(lastDataRow, 1)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(31, 78, 121)
    End With

    ' Borde grueso a la derecha de cada BULK ACTUAL para separar grupos de columnas
    For col = 1 To maxCol
        colC = (col - 1) * 2 + 3   ' columna BULK ACTUAL de cada campo
        With wsC.Range(wsC.Cells(1, colC), wsC.Cells(lastDataRow, colC)).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(31, 78, 121)
        End With
    Next col

    ' Autoajuste de anchos: min 8 para que no queden demasiado estrechas,
    ' max 40 para que no se disparen con textos largos
    wsC.Cells.EntireColumn.AutoFit
    Dim cIdx As Long
    For cIdx = 1 To lastDataCol
        With wsC.Columns(cIdx)
            If .ColumnWidth < 8  Then .ColumnWidth = 8
            If .ColumnWidth > 40 Then .ColumnWidth = 40
        End With
    Next cIdx

    ' AutoFilter en fila 2 (no en fila 1 porque col A tiene celdas fusionadas
    ' y Excel da error si se intenta AutoFilter sobre un rango fusionado)
    wsC.Rows(2).AutoFilter

    ' Congelar filas 1-2 y columna A seleccionando B3
    ' Todo lo que queda arriba y a la izquierda de B3 queda fijo
    Application.ScreenUpdating = True
    wsC.Activate
    wsC.Range("B3").Select
    ActiveWindow.FreezePanes = True
    wsC.Range("A1").Select

    MsgBox "Comparacion completada." & vbCrLf & vbCrLf & _
           "  Registros totales : " & nAll & vbCrLf & _
           "  DIFERENTES        : " & cntSI & vbCrLf & _
           "  IGUALES           : " & cntIguales & vbCrLf & _
           "  SOLO EN V1        : " & cntSoloV1 & vbCrLf & _
           "  SOLO EN V2        : " & cntSoloV2 & vbCrLf & vbCrLf & _
           "Filtra columna FILTRO para ver el resultado por fila.", _
           vbInformation, "Resultado"
    Exit Sub

CompararError:
    ' Algo fallo: restauramos el estado de Excel para no dejarlo bloqueado
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Err.Number <> 0 Then
        MsgBox "Error en la comparacion: " & Err.Description, vbCritical, "Error"
    End If
End Sub


' ================================================================
' INFORME
'
' Genera 3 hojas a partir de la hoja COMPARACION:
'
'   TOTAL        - copia de COMPARACION con todos los registros
'                  y doble columna BULK ANTERIOR/ACTUAL
'
'   DIFERENTES_V2 - solo los registros DIFERENTES y los SOLO EN V2
'                   en formato limpio de una columna por campo
'                   (listo para reimportar o usar directamente)
'
'   SOLO_V1      - solo los registros que estaban en V1 y ya no
'                  estan en V2, en formato limpio
'
' Para las hojas DIFERENTES_V2 y SOLO_V1 se coge el NumberFormat
' de la primera fila de datos de ws2 (no de la cabecera, que
' siempre devuelve "General" y no sirve).
'
' Requiere haber ejecutado CompararHojas antes.
' Tambien necesita que la hoja HOY 2 (v2) siga importada,
' para conocer el formato original de las columnas.
' ================================================================
Sub Informe()

    ' Handler para restaurar Excel si algo peta a mitad
    On Error GoTo InformeError

    ' --- Paso 1: comprobar que COMPARACION existe y tiene datos ---
    Dim wsC As Worksheet
    On Error Resume Next
    Set wsC = ThisWorkbook.Worksheets("COMPARACION")
    On Error GoTo InformeError
    If wsC Is Nothing Then
        MsgBox "Primero ejecuta COMPARAR.", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long, lastCol As Long
    lastRow = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
    lastCol = wsC.Cells(2, wsC.Columns.Count).End(xlToLeft).Column
    If lastRow < 3 Then
        MsgBox "La hoja COMPARACION esta vacia.", vbExclamation
        Exit Sub
    End If

    ' --- Paso 2: necesitamos ws2 para el formato original de columnas ---
    ' Sin ws2 no sabemos el NumberFormat de cada campo
    Dim wsMenu As Worksheet
    Set wsMenu = ThisWorkbook.Worksheets("MENU")
    Dim nomH2 As String
    nomH2 = Trim(wsMenu.Range("J2").Value)
    Dim ws2 As Worksheet
    On Error Resume Next
    Set ws2 = ThisWorkbook.Worksheets(nomH2)
    On Error GoTo InformeError
    If ws2 Is Nothing Then
        MsgBox "No encuentro la hoja HOY 2 (" & nomH2 & ")." & vbCrLf & _
               "Es necesaria para saber el formato original.", vbExclamation
        Exit Sub
    End If

    ' Numero de columnas originales (sin la doble columna de COMPARACION)
    Dim nColsOrig As Long
    nColsOrig = ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column

    ' Buscamos la primera fila de datos en ws2 para el NumberFormat
    ' (la fila de cabecera siempre devuelve "General", no sirve)
    Dim filaDatosWs2 As Long
    filaDatosWs2 = 2
    Dim tmpRow As Long
    For tmpRow = 2 To ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
        If CStr(ws2.Cells(tmpRow, 1).Value) <> "" Then
            filaDatosWs2 = tmpRow
            Exit For
        End If
    Next tmpRow

    ' Apagamos pantalla y alertas para trabajar sin parpadeos
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' ================================================================
    ' HOJA TOTAL
    ' Copia exacta de COMPARACION con todos los registros y la
    ' doble columna BULK ANTERIOR/ACTUAL. Mismo formato y colores.
    ' ================================================================
    On Error Resume Next
    ThisWorkbook.Worksheets("TOTAL").Delete   ' borramos la anterior si existe
    On Error GoTo InformeError

    Dim wsTotal As Worksheet
    Set wsTotal = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsTotal.Name = "TOTAL"

    ' Copiamos cabeceras (filas 1 y 2) y todos los datos
    wsC.Rows(1).Copy wsTotal.Rows(1)
    wsC.Rows(2).Copy wsTotal.Rows(2)
    Dim fOrig As Long, fDest As Long
    fDest = 3
    For fOrig = 3 To lastRow
        wsC.Rows(fOrig).Copy wsTotal.Rows(fDest)
        fDest = fDest + 1
    Next fOrig

    ' Autofit y limites de ancho
    wsTotal.Cells.EntireColumn.AutoFit
    Dim ci As Long
    For ci = 1 To lastCol
        With wsTotal.Columns(ci)
            If .ColumnWidth < 8  Then .ColumnWidth = 8
            If .ColumnWidth > 40 Then .ColumnWidth = 40
        End With
    Next ci
    wsTotal.Rows(2).AutoFilter
    wsTotal.Tab.Color = RGB(31, 78, 121)

    ' FreezePanes necesita la hoja activa y ScreenUpdating = True
    Application.ScreenUpdating = True
    wsTotal.Activate
    wsTotal.Range("B3").Select
    ActiveWindow.FreezePanes = True
    wsTotal.Range("A1").Select
    Application.ScreenUpdating = False

    ' ================================================================
    ' HOJA DIFERENTES_V2
    ' Solo los registros DIFERENTES y los SOLO EN V2.
    ' Formato limpio de una columna por campo, igual que el fichero
    ' original. Se usa el valor de BULK ACTUAL para cada campo.
    ' Esta hoja es la que se exporta con el boton Exportar.
    ' ================================================================
    On Error Resume Next
    ThisWorkbook.Worksheets("DIFERENTES_V2").Delete
    On Error GoTo InformeError

    Dim wsDif As Worksheet
    Set wsDif = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsDif.Name = "DIFERENTES_V2"

    ' Cabecera original de ws2 (mismos campos, mismo orden)
    ws2.Rows(1).Copy wsDif.Rows(1)

    ' Recorremos COMPARACION y copiamos solo las filas que nos interesan
    ' Para cada campo copiamos BULK ACTUAL (colC+1) que es el valor nuevo
    Dim colC    As Long
    Dim c       As Long
    Dim filtro  As String
    Dim celDest As Range
    fDest = 2
    For fOrig = 3 To lastRow
        filtro = CStr(wsC.Cells(fOrig, 1).Value)
        If filtro = "DIFERENTES" Or filtro = "SOLO EN V2" Then
            For c = 1 To nColsOrig
                colC = (c - 1) * 2 + 2   ' columna BULK ANTERIOR en COMPARACION
                Set celDest = wsDif.Cells(fDest, c)
                celDest.Value = wsC.Cells(fOrig, colC + 1).Value   ' BULK ACTUAL
                ' Preservamos el formato de numero/fecha del fichero original
                celDest.NumberFormat = ws2.Cells(filaDatosWs2, c).NumberFormat
            Next c
            fDest = fDest + 1
        End If
    Next fOrig

    wsDif.Cells.EntireColumn.AutoFit
    For ci = 1 To nColsOrig
        With wsDif.Columns(ci)
            If .ColumnWidth < 8  Then .ColumnWidth = 8
            If .ColumnWidth > 40 Then .ColumnWidth = 40
        End With
    Next ci
    With wsDif.Rows(1)   ' cabecera azul como el original
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .RowHeight = 18
    End With
    wsDif.Rows(1).AutoFilter
    wsDif.Tab.Color = RGB(192, 57, 43)

    Application.ScreenUpdating = True
    wsDif.Activate
    wsDif.Rows(2).Select
    ActiveWindow.FreezePanes = True
    wsDif.Range("A1").Select
    Application.ScreenUpdating = False

    ' ================================================================
    ' HOJA SOLO_V1
    ' Solo los registros que estaban en V1 y ya no estan en V2.
    ' Formato limpio usando el valor de BULK ANTERIOR (el que habia).
    ' Estos registros NO se exportan con el boton Exportar.
    ' ================================================================
    On Error Resume Next
    ThisWorkbook.Worksheets("SOLO_V1").Delete
    On Error GoTo InformeError

    Dim wsV1 As Worksheet
    Set wsV1 = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsV1.Name = "SOLO_V1"

    ws2.Rows(1).Copy wsV1.Rows(1)   ' misma cabecera que el original

    fDest = 2
    For fOrig = 3 To lastRow
        If CStr(wsC.Cells(fOrig, 1).Value) = "SOLO EN V1" Then
            For c = 1 To nColsOrig
                colC = (c - 1) * 2 + 2   ' columna BULK ANTERIOR
                Set celDest = wsV1.Cells(fDest, c)
                celDest.Value = wsC.Cells(fOrig, colC).Value
                celDest.NumberFormat = ws2.Cells(filaDatosWs2, c).NumberFormat
            Next c
            fDest = fDest + 1
        End If
    Next fOrig

    wsV1.Cells.EntireColumn.AutoFit
    For ci = 1 To nColsOrig
        With wsV1.Columns(ci)
            If .ColumnWidth < 8  Then .ColumnWidth = 8
            If .ColumnWidth > 40 Then .ColumnWidth = 40
        End With
    Next ci
    With wsV1.Rows(1)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .RowHeight = 18
    End With
    wsV1.Rows(1).AutoFilter
    wsV1.Tab.Color = RGB(31, 78, 121)

    Application.ScreenUpdating = True
    wsV1.Activate
    wsV1.Rows(2).Select
    ActiveWindow.FreezePanes = True
    wsV1.Range("A1").Select

    ' Restauramos el estado y volvemos a COMPARACION
    Application.DisplayAlerts = True
    wsC.Activate
    wsC.Range("A1").Select

    ' Contamos para el resumen final
    Dim cntDif  As Long, cntV2 As Long, cntV1 As Long
    cntDif = 0: cntV2 = 0: cntV1 = 0
    For fOrig = 3 To lastRow
        Select Case CStr(wsC.Cells(fOrig, 1).Value)
            Case "DIFERENTES": cntDif = cntDif + 1
            Case "SOLO EN V2": cntV2 = cntV2 + 1
            Case "SOLO EN V1": cntV1 = cntV1 + 1
        End Select
    Next fOrig

    MsgBox "Hojas de informe generadas:" & vbCrLf & vbCrLf & _
           "  TOTAL         : " & (lastRow - 2) & " registros" & vbCrLf & _
           "  DIFERENTES_V2 : " & cntDif & " diferentes + " & cntV2 & " solo en V2" & vbCrLf & _
           "  SOLO_V1       : " & cntV1 & " registros" & vbCrLf & vbCrLf & _
           "Nota: los SOLO EN V1 no se incluyen en DIFERENTES_V2 ni en la exportacion.", _
           vbInformation, "Informe listo"
    Exit Sub

InformeError:
    ' Algo fallo: restauramos el estado de Excel para no dejarlo bloqueado
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Err.Number <> 0 Then
        MsgBox "Error al generar informe: " & Err.Description, vbCritical, "Error"
    End If
End Sub


' ================================================================
' EXPORTAR
'
' Exporta los registros de la hoja DIFERENTES_V2 a un fichero
' .xlsx nuevo en C:\clientes\bulk\.
'
' Versionado automatico: busca el mayor numero de version
' existente en la carpeta (BULK V01_..., BULK V02_...) y
' propone el siguiente. El usuario puede cambiar el nombre.
'
' Las carpetas C:\clientes y C:\clientes\bulk se crean
' automaticamente si no existen.
'
' El fichero se guarda como .xlsx sin macros (xlOpenXMLWorkbook)
' para que pueda usarse en cualquier sitio sin alertas de seguridad.
'
' Los registros SOLO EN V1 NO se exportan (ya no existen en el
' fichero actual, no tiene sentido incluirlos en la salida).
' ================================================================
Sub Exportar()

    ' --- Paso 1: comprobar que DIFERENTES_V2 existe y tiene datos ---
    Dim wsDif As Worksheet
    On Error Resume Next
    Set wsDif = ThisWorkbook.Worksheets("DIFERENTES_V2")
    On Error GoTo 0
    If wsDif Is Nothing Then
        MsgBox "Primero ejecuta INFORME para generar la hoja DIFERENTES_V2.", _
               vbExclamation, "Falta hoja"
        Exit Sub
    End If

    Dim lastRowDif As Long
    lastRowDif = wsDif.Cells(wsDif.Rows.Count, 1).End(xlUp).Row
    If lastRowDif < 2 Then
        MsgBox "La hoja DIFERENTES_V2 no tiene registros para exportar.", _
               vbExclamation, "Sin datos"
        Exit Sub
    End If

    ' --- Paso 2: crear carpetas si no existen ---
    Dim rutaBase  As String: rutaBase  = "C:\clientes"
    Dim rutaBulk  As String: rutaBulk  = "C:\clientes\bulk"

    If Dir(rutaBase, vbDirectory) = "" Then MkDir rutaBase
    If Dir(rutaBulk, vbDirectory) = "" Then MkDir rutaBulk

    ' --- Paso 3: calcular el siguiente numero de version ---
    ' Buscamos todos los ficheros BULK V*.xlsx en la carpeta
    ' y nos quedamos con el mayor numero encontrado
    Dim maxVer As Integer: maxVer = 0
    Dim f      As String: f = Dir(rutaBulk & "\BULK V*.xlsx")
    Dim posV   As Integer
    Dim posUnd As Integer
    Dim verStr As String
    Do While f <> ""
        posV = InStr(f, " V")
        If posV > 0 Then
            posUnd = InStr(posV, f, "_")
            If posUnd = 0 Then posUnd = Len(f) + 1
            verStr = Mid(f, posV + 2, posUnd - posV - 2)
            If IsNumeric(verStr) Then
                If CInt(verStr) > maxVer Then maxVer = CInt(verStr)
            End If
        End If
        f = Dir
    Loop
    Dim nuevaVer As Integer: nuevaVer = maxVer + 1

    ' --- Paso 4: sugerir nombre con version y fecha/hora ---
    ' Formato: BULK V01_130126_1430
    Dim fechaHora   As String: fechaHora   = Format(Now, "ddmmyy_hhmm")
    Dim verStr2     As String: verStr2     = Right("0" & CStr(nuevaVer), 2)
    Dim nomSugerido As String: nomSugerido = "BULK V" & verStr2 & "_" & fechaHora

    ' Application.InputBox con Type:=2 devuelve False al cancelar (no "")
    ' y String con el valor si el usuario acepta
    Dim nomFinalVar As Variant
    nomFinalVar = Application.InputBox("Nombre del fichero a exportar:" & vbCrLf & _
                                       "(se guardara en " & rutaBulk & "\)", _
                                       "Exportar", Default:=nomSugerido, Type:=2)

    If VarType(nomFinalVar) = vbBoolean Then Exit Sub   ' pulso Cancelar

    Dim nomFinal As String
    nomFinal = CStr(nomFinalVar)
    If Trim(nomFinal) = "" Then
        MsgBox "El nombre no puede estar vacio.", vbExclamation
        Exit Sub
    End If

    ' Quitamos .xlsx si lo escribio a mano para no duplicar la extension
    If LCase(Right(nomFinal, 5)) = ".xlsx" Then
        nomFinal = Left(nomFinal, Len(nomFinal) - 5)
    End If

    Dim rutaFinal As String
    rutaFinal = rutaBulk & "\" & nomFinal & ".xlsx"

    ' --- Paso 5: confirmar si el fichero ya existe ---
    If Dir(rutaFinal) <> "" Then
        Dim overwrite As Integer
        overwrite = MsgBox("Ya existe el fichero:" & vbCrLf & rutaFinal & vbCrLf & vbCrLf & _
                           "Sobreescribir?", vbQuestion + vbYesNo)
        If overwrite = vbNo Then Exit Sub
    End If

    ' --- Paso 6: crear el libro nuevo y copiar la hoja ---
    Application.ScreenUpdating = False

    Dim wbNuevo As Workbook
    Set wbNuevo = Workbooks.Add

    ' Copiamos DIFERENTES_V2 antes de la primera hoja del libro nuevo
    wsDif.Copy Before:=wbNuevo.Worksheets(1)

    ' Borramos las hojas vacias que Excel añade al crear el libro
    ' En orden inverso para no saltarnos ninguna
    Application.DisplayAlerts = False
    Dim iExtra   As Integer
    Dim nomExtra As String
    For iExtra = wbNuevo.Worksheets.Count To 1 Step -1
        nomExtra = wbNuevo.Worksheets(iExtra).Name
        If nomExtra <> "DIFERENTES_V2" Then
            wbNuevo.Worksheets(iExtra).Delete
        End If
    Next iExtra
    Application.DisplayAlerts = True

    ' Renombramos la hoja copiada a "BULK" para que quede limpia
    Dim wsRename As Worksheet
    For Each wsRename In wbNuevo.Worksheets
        If wsRename.Name = "DIFERENTES_V2" Then
            wsRename.Name = "BULK"
            Exit For
        End If
    Next wsRename

    ' Guardamos como xlsx sin macros y cerramos
    wbNuevo.SaveAs Filename:=rutaFinal, FileFormat:=xlOpenXMLWorkbook
    wbNuevo.Close SaveChanges:=False

    Application.ScreenUpdating = True

    MsgBox "Fichero exportado correctamente:" & vbCrLf & vbCrLf & _
           "  " & rutaFinal & vbCrLf & vbCrLf & _
           "  Registros exportados : " & (lastRowDif - 1) & vbCrLf & vbCrLf & _
           "Nota: los registros SOLO EN V1 no se han exportado.", _
           vbInformation, "Exportacion completada"
End Sub
