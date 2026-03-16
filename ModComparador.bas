' ============================================================
' COMPARADOR DE EXCELS
' ============================================================
' Esto lo montamos para comparar dos ficheros de empleados
' que tienen la misma estructura pero el contenido puede cambiar
' entre una semana y otra.
'
' La clave de cruce es Employee ID - no comparamos por posicion
' de fila sino buscando el mismo empleado en los dos ficheros.
'
' Cada campo aparece dos veces en el resultado: PAGE1 y PAGE2,
' para ver de un vistazo que habia antes y que hay ahora.
'
' Botones del MENU:
'   - IMPORTAR HOY 1  -> coge el primer fichero abierto que elijas
'   - IMPORTAR HOY 2  -> coge el segundo
'   - COMPARAR        -> genera la hoja COMPARACION
'   - BORRAR TODO     -> limpia todo y deja solo el MENU
' ============================================================


' ------------------------------------------------------------
' Estos dos son los botones del menu - solo llaman a ImportarHoja
' con el numero de slot correspondiente (1 o 2)
' ------------------------------------------------------------
Sub ImportarHoy1()
    ImportarHoja 1
End Sub

Sub ImportarHoy2()
    ImportarHoja 2
End Sub


' ============================================================
' BORRAR TODO
' Elimina todas las pestanas excepto MENU y limpia las
' referencias guardadas en J1 y J2.
' Util para empezar de cero sin cerrar el fichero.
' ============================================================
Sub BorrarTodo()

    ' Preguntamos antes de borrar, no sea que se pulse por error
    Dim resp As Integer
    resp = MsgBox("Se eliminaran todas las hojas excepto MENU." & vbCrLf & _
                  "Esto incluye las importadas y la comparacion." & vbCrLf & vbCrLf & _
                  "Continuar?", vbQuestion + vbYesNo, "Confirmar borrado")
    If resp = vbNo Then Exit Sub

    ' Cogemos referencia al MENU antes de ponernos a borrar
    Dim wsMenu As Worksheet
    Set wsMenu = ThisWorkbook.Worksheets("MENU")

    ' Limpiamos las celdas ocultas donde guardamos los nombres
    ' de las hojas importadas (J1 = PAGE1, J2 = PAGE2)
    wsMenu.Range("J1").Value = ""
    wsMenu.Range("J2").Value = ""

    ' Recorremos todas las hojas y borramos las que no sean MENU
    ' DisplayAlerts = False para que no pregunte por cada una
    Application.DisplayAlerts = False
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "MENU" Then ws.Delete
    Next ws
    Application.DisplayAlerts = True

    wsMenu.Activate
    MsgBox "Listo. Solo queda la hoja MENU.", vbInformation, "Borrado completado"
End Sub


' ============================================================
' IMPORTAR HOJA
' Recibe slot = 1 o 2 segun que boton se pulsó.
' Busca los ficheros Excel abiertos (excepto este),
' pregunta cual quieres usar, coge siempre la primera hoja
' (que es PAGE 1 en nuestro caso) y la copia aqui dentro
' con el nombre original + sufijo v1 o v2.
' ============================================================
Sub ImportarHoja(slot As Integer)

    ' Guardamos referencia al MENU al principio del todo.
    ' Es importante hacerlo ANTES de cualquier operacion Copy
    ' porque despues Excel puede cambiar el libro activo y
    ' ThisWorkbook puede apuntar al sitio equivocado.
    Dim wsMenu As Worksheet
    Set wsMenu = ThisWorkbook.Worksheets("MENU")

    ' ----------------------------------------------------------
    ' PASO 1: miramos que ficheros hay abiertos ademas de este
    ' Los metemos en un array para poder mostrarlos numerados
    ' ----------------------------------------------------------
    Dim nombres() As String
    Dim wbLoop    As Workbook
    Dim n         As Integer
    Dim i         As Integer
    n = 0

    For Each wbLoop In Application.Workbooks
        ' Saltamos este mismo fichero, solo nos interesan los externos
        If wbLoop.Name <> ThisWorkbook.Name Then
            ReDim Preserve nombres(n)
            nombres(n) = wbLoop.Name
            n = n + 1
        End If
    Next wbLoop

    ' Si no hay nada abierto avisamos y salimos
    If n = 0 Then
        MsgBox "No hay otros ficheros Excel abiertos." & vbCrLf & _
               "Abre primero el fichero que quieres importar.", _
               vbExclamation, "Sin ficheros"
        Exit Sub
    End If

    ' ----------------------------------------------------------
    ' PASO 2: mostramos la lista y pedimos que elijan uno
    ' Usamos InputBox con Type:=1 para forzar numero
    ' Si pulsan Cancelar, VarType devuelve vbBoolean -> salimos
    ' ----------------------------------------------------------
    Dim lista As String
    lista = "Ficheros Excel abiertos:" & vbCrLf & vbCrLf
    For i = 0 To n - 1
        lista = lista & "  " & (i + 1) & "  ->  " & nombres(i) & vbCrLf
    Next i
    lista = lista & vbCrLf & "Escribe el numero del fichero:"

    Dim respWB As Variant
    respWB = Application.InputBox(lista, "Importar HOY " & slot, Type:=1)

    ' Cancelar devuelve False (booleano), no una cadena "False"
    If VarType(respWB) = vbBoolean Then Exit Sub

    ' Validamos que el numero este en rango
    Dim idxWB As Integer
    idxWB = CInt(respWB) - 1
    If idxWB < 0 Or idxWB >= n Then
        MsgBox "Numero fuera de rango (1 a " & n & ").", vbExclamation
        Exit Sub
    End If

    ' Ya tenemos el workbook origen
    Dim wbOrigen As Workbook
    Set wbOrigen = Application.Workbooks(nombres(idxWB))

    ' ----------------------------------------------------------
    ' PASO 3: cogemos siempre la primera hoja del fichero
    ' En nuestro caso siempre se llama PAGE 1 y es la unica
    ' que nos interesa, asi que no hace falta preguntar
    ' ----------------------------------------------------------
    Dim wsOrigen As Worksheet
    Set wsOrigen = wbOrigen.Worksheets(1)

    ' ----------------------------------------------------------
    ' PASO 4: construimos el nombre que tendra la pestana aqui
    ' Usamos el nombre de la hoja origen + sufijo " v1" o " v2"
    ' Si el nombre es muy largo lo recortamos (Excel tiene limite)
    ' Tambien limpiamos caracteres que Excel no permite en nombres
    ' ----------------------------------------------------------
    Dim nomBase As String
    nomBase = wsOrigen.Name
    If Len(nomBase) > 25 Then nomBase = Left(nomBase, 25)

    Dim nomHoja As String
    nomHoja = nomBase & " v" & slot

    ' Estos caracteres no pueden ir en nombres de hoja en Excel
    Dim cars As Variant
    Dim c    As Variant
    cars = Array("/", "\", "?", "*", "[", "]", ":")
    For Each c In cars
        nomHoja = Replace(nomHoja, CStr(c), "_")
    Next c

    ' ----------------------------------------------------------
    ' PASO 5: si ya existia una hoja con ese nombre la borramos
    ' para sustituirla con la nueva importacion
    ' ----------------------------------------------------------
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(nomHoja).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' ----------------------------------------------------------
    ' PASO 6: copiamos la hoja al final de este libro
    ' Guardamos wbThis y el numero de hojas ANTES del Copy
    ' porque despues el libro activo puede cambiar y perder
    ' la referencia a ThisWorkbook
    ' ----------------------------------------------------------
    Dim wbThis     As Workbook
    Dim totalHojas As Integer
    Set wbThis = ThisWorkbook
    totalHojas = wbThis.Worksheets.Count

    wsOrigen.Copy After:=wbThis.Worksheets(totalHojas)

    ' La hoja copiada siempre queda al final, le ponemos el nombre
    wbThis.Worksheets(wbThis.Worksheets.Count).Name = nomHoja

    ' ----------------------------------------------------------
    ' PASO 7: guardamos el nombre en una celda oculta del MENU
    ' J1 = nombre de la hoja PAGE1, J2 = nombre de la hoja PAGE2
    ' Asi la macro de comparacion sabe donde buscar
    ' ----------------------------------------------------------
    If slot = 1 Then
        wsMenu.Range("J1").Value = nomHoja
    Else
        wsMenu.Range("J2").Value = nomHoja
    End If

    ' Volvemos al MENU sin avisar, mas limpio
    wsMenu.Activate

End Sub


' ============================================================
' COMPARAR HOJAS
'
' Esta es la macro principal. Hace lo siguiente:
'
'  1. Lee los nombres de las hojas importadas desde MENU J1/J2
'  2. Busca la columna Employee ID en cada hoja
'  3. Construye la lista completa de IDs (union de los dos ficheros)
'  4. Crea la hoja COMPARACION con doble columna por campo (PAGE1/PAGE2)
'  5. Para cada empleado busca su fila en cada fichero y escribe
'     los valores lado a lado
'  6. Segun el resultado marca la fila con color y escribe en
'     la columna DIFERENTE: IGUALES / DIFERENTES / SOLO EN V1 / SOLO EN V2
'  7. Aplica bordes, autoajusta anchos y activa el autofiltro
'
' Colores de la columna DIFERENTE:
'   IGUALES    -> verde oscuro   (sin cambios)
'   DIFERENTES -> rojo oscuro    (hay campos que han cambiado)
'   SOLO EN V1 -> azul oscuro    (estaba en el fichero antiguo pero ya no esta)
'   SOLO EN V2 -> verde oscuro   (es una alta nueva, no estaba antes)
'
' Para SOLO EN V1 ademas se tacha el texto de la fila (ha desaparecido)
' Para SOLO EN V2 el fondo de la fila es verde suave (es nuevo)
' ============================================================
Sub CompararHojas()

    ' Referencia al MENU capturada al inicio, antes de todo
    Dim wsMenu As Worksheet
    Set wsMenu = ThisWorkbook.Worksheets("MENU")

    ' Leemos los nombres de las hojas que se importaron
    ' Estan guardados en J1 y J2 por la macro ImportarHoja
    Dim nomH1 As String
    Dim nomH2 As String
    nomH1 = Trim(wsMenu.Range("J1").Value)
    nomH2 = Trim(wsMenu.Range("J2").Value)

    ' Si alguna esta vacia es que no se ha importado todavia
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
    On Error GoTo 0

    ' Si no las encuentra avisamos con el nombre exacto para ayudar a depurar
    If ws1 Is Nothing Then
        MsgBox "No encuentro la hoja HOY 1:" & vbCrLf & nomH1, vbCritical
        Exit Sub
    End If
    If ws2 Is Nothing Then
        MsgBox "No encuentro la hoja HOY 2:" & vbCrLf & nomH2, vbCritical
        Exit Sub
    End If

    ' ----------------------------------------------------------
    ' Calculamos hasta donde llegan los datos en cada hoja
    ' Buscamos la ultima fila con datos en columna A
    ' y la ultima columna con datos en fila 1 (cabeceras)
    ' ----------------------------------------------------------
    Dim lastRow1 As Long, lastCol1 As Long
    Dim lastRow2 As Long, lastCol2 As Long
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastCol1 = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    lastCol2 = ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column

    ' El maximo de columnas es el mayor de los dos ficheros
    ' por si alguno tiene columnas extra que el otro no tiene
    Dim maxCol As Long
    maxCol = Application.Max(lastCol1, lastCol2)

    ' ----------------------------------------------------------
    ' Buscamos la columna Employee ID en cada hoja
    ' Comparamos en minusculas y sin espacios para ser flexibles
    ' ----------------------------------------------------------
    Dim colEmp1 As Long
    Dim colEmp2 As Long
    Dim col     As Long
    colEmp1 = 0
    colEmp2 = 0

    For col = 1 To lastCol1
        If LCase(Trim(CStr(ws1.Cells(1, col).Value))) = "* employee id" Then
            colEmp1 = col
            Exit For    ' ya lo encontramos, no seguimos buscando
        End If
    Next col

    For col = 1 To lastCol2
        If LCase(Trim(CStr(ws2.Cells(1, col).Value))) = "* employee id" Then
            colEmp2 = col
            Exit For
        End If
    Next col

    ' Si no encontramos Employee ID en alguno de los dos avisamos
    If colEmp1 = 0 Or colEmp2 = 0 Then
        MsgBox "No se encuentra '* Employee ID' en " & _
               IIf(colEmp1 = 0, nomH1, nomH2) & "." & vbCrLf & _
               "Verifica que el nombre es exactamente '* Employee ID'.", _
               vbExclamation, "Columna no encontrada"
        Exit Sub
    End If

    ' ----------------------------------------------------------
    ' Construimos un indice de ws2 para buscar rapido por ID
    ' Dos arrays paralelos: uno con los IDs, otro con el numero
    ' de fila donde esta ese ID en ws2.
    ' Asi en el bucle principal solo hacemos un recorrido simple
    ' en lugar de buscar en toda la hoja cada vez.
    ' ----------------------------------------------------------
    Dim idx2Keys() As String
    Dim idx2Rows() As Long
    Dim nIdx2      As Long
    nIdx2 = lastRow2 - 1   ' filas de datos (sin cabecera)
    ReDim idx2Keys(1 To nIdx2)
    ReDim idx2Rows(1 To nIdx2)
    Dim e As Long
    For e = 1 To nIdx2
        idx2Keys(e) = CStr(ws2.Cells(e + 1, colEmp2).Value)
        idx2Rows(e) = e + 1   ' fila real en la hoja (e+1 porque fila 1 es cabecera)
    Next e

    ' ----------------------------------------------------------
    ' Construimos la lista completa de Employee IDs
    ' Es la union de los dos ficheros: primero los de ws1,
    ' luego los de ws2 que no estaban ya en la lista.
    ' Asi no se nos escapa nadie aunque solo aparezca en uno.
    ' ----------------------------------------------------------
    Dim allIDs() As String
    Dim nAll     As Long
    Dim fila     As Long
    Dim found    As Boolean
    Dim k        As Long
    nAll = 0

    ' Recorremos ws1 y vamos metiendo IDs unicos
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

    ' Ahora los de ws2 que no estuvieran ya en la lista
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

    ' Ordenamos los IDs para que el resultado salga ordenado
    ' Usamos burbuja simple - para listas de empleados va sobrado
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

    ' ----------------------------------------------------------
    ' Borramos la hoja COMPARACION si ya existia de antes
    ' para generarla limpia cada vez
    ' ----------------------------------------------------------
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("COMPARACION").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Creamos la nueva hoja COMPARACION al final del libro
    Dim wsC As Worksheet
    Set wsC = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsC.Name = "COMPARACION"

    ' ----------------------------------------------------------
    ' CABECERAS - dos filas:
    '   Fila 1: nombre del campo, fusionado sobre PAGE1 y PAGE2
    '   Fila 2: etiquetas PAGE1 / PAGE2
    '
    ' La columna FILTRO va en primer lugar (columna 1).
    ' Los campos de datos empiezan en columna 2.
    ' La formula para cada campo es:
    '   colC = (col-1)*2 + 2  -> columna PAGE1 del campo col
    '   colC+1                -> columna PAGE2 del campo col
    ' ----------------------------------------------------------
    Dim colC     As Long
    Dim cabecera As String

    ' Columna FILTRO primero: ocupa columna 1, filas 1 y 2
    ' La fusionamos en vertical para que ocupe las dos filas de cabecera
    wsC.Range(wsC.Cells(1, 1), wsC.Cells(2, 1)).Merge
    wsC.Cells(1, 1).Value = "FILTRO"

    ' Ahora los campos de datos, desplazados una columna a la derecha
    For col = 1 To maxCol
        colC = (col - 1) * 2 + 2   ' PAGE1 de este campo (empieza en col 2)

        ' Cogemos el nombre del campo de ws1 si existe ahi,
        ' si no de ws2 (por si el campo es nuevo en la segunda version)
        If col <= lastCol1 Then
            cabecera = CStr(ws1.Cells(1, col).Value)
        ElseIf col <= lastCol2 Then
            cabecera = CStr(ws2.Cells(1, col).Value)
        Else
            cabecera = "Campo" & col
        End If

        ' Fusionamos PAGE1 y PAGE2 en fila 1 para el nombre del campo
        wsC.Range(wsC.Cells(1, colC), wsC.Cells(1, colC + 1)).Merge
        wsC.Cells(1, colC).Value = cabecera

        ' En fila 2 ponemos PAGE1 y PAGE2
        wsC.Cells(2, colC).Value = "PAGE1"
        wsC.Cells(2, colC + 1).Value = "PAGE2"
    Next col

    ' Formato fila 1: azul oscuro, texto blanco, negrita
    With wsC.Rows(1)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 20
    End With

    ' Formato fila 2: azul medio, texto blanco, negrita
    With wsC.Rows(2)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(41, 128, 185)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 18
    End With

    ' La columna FILTRO (col 1) la mantenemos azul oscuro en ambas filas
    ' para que destaque visualmente del resto de cabeceras
    wsC.Cells(1, 1).Interior.Color = RGB(31, 78, 121)
    wsC.Cells(2, 1).Interior.Color = RGB(31, 78, 121)

    ' ----------------------------------------------------------
    ' BUCLE PRINCIPAL: recorremos todos los Employee IDs
    ' y para cada uno buscamos su fila en ws1 y en ws2.
    '
    ' Hay cuatro situaciones posibles:
    '   r1>0 y r2>0  -> el empleado existe en los dos ficheros
    '   r1=0         -> solo esta en ws2 (alta nueva)
    '   r2=0         -> solo esta en ws1 (ya no esta en el nuevo)
    '   (r1=0 y r2=0 no puede pasar porque el ID vino de alguno)
    ' ----------------------------------------------------------
    Application.ScreenUpdating = False   ' apagamos la pantalla para ir mas rapido

    Dim filaC      As Long    ' fila actual en wsC donde escribimos
    Dim r1         As Long    ' fila del empleado en ws1 (0 si no esta)
    Dim r2         As Long    ' fila del empleado en ws2 (0 si no esta)
    Dim v1         As String  ' valor de la celda en ws1
    Dim v2         As String  ' valor de la celda en ws2
    Dim difFila    As Boolean ' hay alguna diferencia en esta fila?
    Dim estado     As String  ' IGUALES / SI / SOLO EN V1 / SOLO EN V2

    ' Contadores para el resumen final.
    ' Los acumulamos en el bucle porque es mas fiable que CountIf
    ' cuando los valores de la columna FILTRO pueden variar.
    Dim cntSI      As Long    ' filas con campos diferentes
    Dim cntSoloV1  As Long    ' filas que solo estan en PAGE1
    Dim cntSoloV2  As Long    ' filas que solo estan en PAGE2
    Dim cntIguales As Long    ' filas identicas en los dos ficheros

    ' Ultima columna de datos (para colorear solo hasta ahi, no FILTRO)
    Dim lastDataCol As Long
    lastDataCol = maxCol * 2 + 1   ' ultima columna PAGE2 del ultimo campo

    filaC = 3       ' los datos empiezan en fila 3 (filas 1 y 2 son cabeceras)
    cntSI      = 0
    cntSoloV1  = 0
    cntSoloV2  = 0
    cntIguales = 0

    For k = 1 To nAll
        Dim empID As String
        empID = allIDs(k)   ' ID del empleado que estamos procesando ahora

        ' Buscamos la fila de este empleado en ws1
        r1 = 0
        For fila = 2 To lastRow1
            If CStr(ws1.Cells(fila, colEmp1).Value) = empID Then
                r1 = fila
                Exit For   ' encontrado, salimos del bucle
            End If
        Next fila

        ' Buscamos la fila de este empleado en ws2 usando el indice
        r2 = 0
        For e = 1 To nIdx2
            If idx2Keys(e) = empID Then
                r2 = idx2Rows(e)
                Exit For
            End If
        Next e

        ' Determinamos el estado inicial segun donde aparece el empleado
        difFila = False
        If r1 = 0 Then
            ' No esta en ws1 -> es una alta nueva en ws2
            estado = "SOLO EN V2"
            difFila = True
        ElseIf r2 = 0 Then
            ' No esta en ws2 -> estaba en ws1 pero ya no esta
            estado = "SOLO EN V1"
            difFila = True
        Else
            ' Esta en los dos -> empezamos asumiendo que son iguales
            ' y lo cambiaremos a SI si encontramos alguna diferencia
            estado = "NO"
        End If

        ' Escribimos los valores de cada campo lado a lado (PAGE1 / PAGE2)
        ' Si el empleado no esta en uno de los ficheros esa columna queda vacia.
        ' Los campos empiezan en columna 2 porque la columna 1 es FILTRO.
        For col = 1 To maxCol
            colC = (col - 1) * 2 + 2   ' columna PAGE1 de este campo (empieza en 2)
            v1 = ""
            v2 = ""
            If r1 > 0 And col <= lastCol1 Then v1 = CStr(ws1.Cells(r1, col).Value)
            If r2 > 0 And col <= lastCol2 Then v2 = CStr(ws2.Cells(r2, col).Value)
            wsC.Cells(filaC, colC).Value = v1       ' PAGE1
            wsC.Cells(filaC, colC + 1).Value = v2   ' PAGE2

            ' Si encontramos una diferencia cambiamos el estado.
            ' Solo actualizamos si todavia era "NO" para no machacar
            ' un estado SOLO EN V1/V2 ya establecido antes del bucle.
            If estado = "NO" And v1 <> v2 Then
                estado = "SI"
                difFila = True
            End If
        Next col

        ' ----------------------------------------------------------
        ' Marcamos la fila segun el estado que ha resultado.
        '
        ' Colores de fila (solo hasta lastDataCol, no toca la col FILTRO):
        '   IGUALES    -> verde suave en toda la fila de datos
        '   DIFERENTES -> blanco (sin fondo), solo se marcan las celdas PAGE2 distintas
        '   SOLO EN V1 -> azul suave + tachado (ha desaparecido)
        '   SOLO EN V2 -> amarillo suave (es una alta nueva)
        '
        ' La columna FILTRO (col 1) recibe el mismo color oscuro que
        ' el fondo de la fila, para que destaque y sea coherente.
        ' ----------------------------------------------------------
        If estado = "SI" Then
            ' Hay campos diferentes. Fondo blanco en datos, rojo en FILTRO.
            ' Marcamos en rojo oscuro las celdas PAGE2 que cambiaron.
            With wsC.Cells(filaC, 1)   ' columna FILTRO
                .Value = "DIFERENTES"
                .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(192, 57, 43)   ' rojo oscuro
            End With
            ' Recorremos los campos para marcar los que cambiaron en PAGE2
            For col = 1 To maxCol
                colC = (col - 1) * 2 + 2
                If CStr(wsC.Cells(filaC, colC).Value) <> _
                   CStr(wsC.Cells(filaC, colC + 1).Value) Then
                    With wsC.Cells(filaC, colC + 1)
                        .Interior.Color = RGB(139, 0, 0)   ' rojo oscuro en celda PAGE2
                        .Font.Color = RGB(255, 255, 255)
                        .Font.Bold = True
                    End With
                End If
            Next col

        ElseIf estado = "SOLO EN V1" Then
            ' Estaba en PAGE1 pero ya no aparece en PAGE2.
            ' Fondo azul suave + tachado en columnas de datos.
            wsC.Range(wsC.Cells(filaC, 2), wsC.Cells(filaC, lastDataCol)).Interior.Color = RGB(213, 229, 242)
            With wsC.Range(wsC.Cells(filaC, 2), wsC.Cells(filaC, lastDataCol)).Font
                .Strikethrough = True
                .Color = RGB(80, 80, 80)
            End With
            ' Columna FILTRO: azul oscuro, sin tachado, texto blanco
            With wsC.Cells(filaC, 1)
                .Value = "SOLO EN V1"
                .Font.Bold = True
                .Font.Strikethrough = False
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(31, 78, 121)   ' azul oscuro
            End With

        ElseIf estado = "SOLO EN V2" Then
            ' Alta nueva, no estaba en PAGE1.
            ' Fondo amarillo suave en datos, amarillo oscuro en FILTRO.
            wsC.Range(wsC.Cells(filaC, 2), wsC.Cells(filaC, lastDataCol)).Interior.Color = RGB(255, 242, 204)
            With wsC.Cells(filaC, 1)
                .Value = "SOLO EN V2"
                .Font.Bold = True
                .Font.Strikethrough = False
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(156, 101, 0)   ' ambar oscuro
            End With

        Else
            ' IGUALES: el empleado existe en los dos y todos los campos son iguales.
            ' Fondo verde suave en datos, verde oscuro en FILTRO.
            wsC.Range(wsC.Cells(filaC, 2), wsC.Cells(filaC, lastDataCol)).Interior.Color = RGB(198, 239, 206)
            With wsC.Cells(filaC, 1)
                .Value = "IGUALES"
                .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(39, 120, 56)   ' verde oscuro
            End With
        End If

        ' Acumulamos en el contador correspondiente para el resumen final
        ' No podemos usar CountIf porque SOLO EN V1/V2 tienen la misma
        ' celda DIFERENTE con distintos textos y esto es mas seguro
        If estado = "SI" Then
            cntSI = cntSI + 1
        ElseIf estado = "SOLO EN V1" Then
            cntSoloV1 = cntSoloV1 + 1
        ElseIf estado = "SOLO EN V2" Then
            cntSoloV2 = cntSoloV2 + 1
        Else
            cntIguales = cntIguales + 1
        End If

        filaC = filaC + 1   ' siguiente fila del resultado
    Next k

    ' ----------------------------------------------------------
    ' FORMATO FINAL DE LA HOJA COMPARACION
    ' ----------------------------------------------------------

    ' Borde grueso azul a la derecha de la columna FILTRO (col 1)
    ' para separarla visualmente de los datos
    Dim lastDataRow As Long
    lastDataRow = filaC - 1   ' ultima fila con datos

    With wsC.Range(wsC.Cells(1, 1), wsC.Cells(lastDataRow, 1)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(31, 78, 121)
    End With

    ' Borde grueso a la derecha de cada columna PAGE2 de cada campo
    ' para separar visualmente un grupo del siguiente
    For col = 1 To maxCol
        colC = (col - 1) * 2 + 3   ' columna PAGE2 de cada campo (offset 2 por FILTRO + 1 por PAGE1)
        With wsC.Range(wsC.Cells(1, colC), wsC.Cells(lastDataRow, colC)).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(31, 78, 121)
        End With
    Next col

    ' Autoajustamos el ancho de cada columna para que no se recorte nada
    ' Ponemos un minimo de 8 y un maximo de 40 para que no quede
    ' ni demasiado estrecha ni demasiado ancha
    wsC.Cells.EntireColumn.AutoFit
    Dim cIdx As Long
    For cIdx = 1 To colDif
        With wsC.Columns(cIdx)
            If .ColumnWidth < 8  Then .ColumnWidth = 8
            If .ColumnWidth > 40 Then .ColumnWidth = 40
        End With
    Next cIdx

    ' Autofiltro desde celda A1 - Excel lo aplica sobre toda la cabecera
    ' La columna FILTRO al inicio permite filtrar directamente por estado
    wsC.Range("A1").AutoFilter

    ' Volvemos a encender la pantalla antes de mover el foco
    Application.ScreenUpdating = True

    ' Congelamos las dos filas de cabecera para que siempre se vean
    ' al hacer scroll hacia abajo
    wsC.Activate
    wsC.Rows(3).Select
    ActiveWindow.FreezePanes = True
    wsC.Range("A1").Select

    ' ----------------------------------------------------------
    ' RESUMEN: mostramos cuantos empleados hay en cada categoria
    ' Los contadores los fuimos acumulando en el bucle principal
    ' La suma de los 4 debe dar siempre nAll (total de empleados)
    ' ----------------------------------------------------------
    MsgBox "Comparacion completada." & vbCrLf & vbCrLf & _
           "  Registros totales : " & nAll & vbCrLf & _
           "  DIFERENTES        : " & cntSI & vbCrLf & _
           "  IGUALES           : " & cntIguales & vbCrLf & _
           "  SOLO EN V1        : " & cntSoloV1 & vbCrLf & _
           "  SOLO EN V2        : " & cntSoloV2 & vbCrLf & vbCrLf & _
           "Filtra columna FILTRO para ver el resultado por fila.", _
           vbInformation, "Resultado"

End Sub
