' ============================================================
' MODULO INDEPENDIENTE - CREAR BOTONES EN EL MENU
' No toca nada del comparador. Ejecutar una sola vez.
' Si lo vuelves a ejecutar borra los botones anteriores y los rehace.
' ============================================================
Sub CrearBotonesMenu()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("MENU")

    ' Borramos shapes anteriores por si se ejecuta mas de una vez
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    Application.ScreenUpdating = False

    ' Medidas
    Dim btnH  As Double: btnH  = 40
    Dim btnW  As Double: btnW  = 210
    Dim lblH  As Double: lblH  = 20
    Dim gap   As Double: gap   = 12
    Dim leftX As Double: leftX = 40
    Dim topY  As Double: topY  = 65

    ' ----------------------------------------------------------
    ' Titulo
    ' ----------------------------------------------------------
    Dim tit As Shape
    Set tit = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftX, 12, btnW, 38)
    With tit.TextFrame2
        .TextRange.Text = "COMPARADOR DE EXCELS"
        .TextRange.Font.Size = 16
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
    End With
    With tit.Fill: .Solid: .ForeColor.RGB = RGB(31, 78, 121): End With
    tit.Line.Visible = msoFalse

    ' ----------------------------------------------------------
    ' Helper interno: crea label + boton y devuelve el boton
    ' ----------------------------------------------------------

    ' PASO 1 - Importar bulk antiguo
    Dim top1 As Double: top1 = topY
    Dim l1 As Shape
    Set l1 = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftX, top1, btnW, lblH)
    With l1.TextFrame2
        .TextRange.Text = "Paso 1  |  Importar bulk antiguo"
        .TextRange.Font.Size = 8
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    With l1.Fill: .Solid: .ForeColor.RGB = RGB(21, 67, 96): End With
    l1.Line.Visible = msoFalse

    Dim b1 As Shape
    Set b1 = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftX, top1 + lblH, btnW, btnH)
    With b1.TextFrame2
        .TextRange.Text = "IMPORTAR FICHERO 1"
        .TextRange.Font.Size = 11
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
    End With
    With b1.Fill: .Solid: .ForeColor.RGB = RGB(31, 97, 141): End With
    b1.Line.ForeColor.RGB = RGB(21, 67, 96)
    b1.OnAction = "ImportarHoy1"

    ' PASO 2 - Importar bulk actual
    Dim top2 As Double: top2 = top1 + lblH + btnH + gap
    Dim l2 As Shape
    Set l2 = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftX, top2, btnW, lblH)
    With l2.TextFrame2
        .TextRange.Text = "Paso 2  |  Importar bulk actual"
        .TextRange.Font.Size = 8
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    With l2.Fill: .Solid: .ForeColor.RGB = RGB(11, 83, 69): End With
    l2.Line.Visible = msoFalse

    Dim b2 As Shape
    Set b2 = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftX, top2 + lblH, btnW, btnH)
    With b2.TextFrame2
        .TextRange.Text = "IMPORTAR FICHERO 2"
        .TextRange.Font.Size = 11
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
    End With
    With b2.Fill: .Solid: .ForeColor.RGB = RGB(17, 122, 101): End With
    b2.Line.ForeColor.RGB = RGB(11, 83, 69)
    b2.OnAction = "ImportarHoy2"

    ' PASO 3 - Comparar
    Dim top3 As Double: top3 = top2 + lblH + btnH + gap
    Dim l3 As Shape
    Set l3 = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftX, top3, btnW, lblH)
    With l3.TextFrame2
        .TextRange.Text = "Paso 3  |  Comparar"
        .TextRange.Font.Size = 8
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    With l3.Fill: .Solid: .ForeColor.RGB = RGB(120, 40, 31): End With
    l3.Line.Visible = msoFalse

    Dim b3 As Shape
    Set b3 = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftX, top3 + lblH, btnW, btnH)
    With b3.TextFrame2
        .TextRange.Text = "COMPARAR"
        .TextRange.Font.Size = 11
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
    End With
    With b3.Fill: .Solid: .ForeColor.RGB = RGB(192, 57, 43): End With
    b3.Line.ForeColor.RGB = RGB(120, 40, 31)
    b3.OnAction = "CompararHojas"

    ' BORRAR - separado, sin label de paso
    Dim top4 As Double: top4 = top3 + lblH + btnH + gap * 3
    Dim b4 As Shape
    Set b4 = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftX, top4, btnW, btnH)
    With b4.TextFrame2
        .TextRange.Text = "BORRAR TODAS LAS HOJAS"
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(180, 180, 180)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
    End With
    With b4.Fill: .Solid: .ForeColor.RGB = RGB(60, 60, 60): End With
    b4.Line.ForeColor.RGB = RGB(40, 40, 40)
    b4.OnAction = "BorrarTodo"

    ' Color de la pestana MENU
    ws.Tab.Color = RGB(31, 78, 121)

    Application.ScreenUpdating = True
    MsgBox "Botones creados.", vbInformation, "Listo"

End Sub
