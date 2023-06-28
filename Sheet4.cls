VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Const ListN3 = "СХЕМА"
Const ListN4 = "ДАННЫЕ"

Private Sub CommandButton1_Click()
    RenewColorList (ComboBox1.Value)
    Call ReDraw(ComboBox1.Value)
    MsgBox "Обновление завершено"
End Sub

Private Sub CommandButton2_Click()
    CreateLegend (ComboBox1.Value)
End Sub


'Поиск элемента в таблице. Если Элемент найден, возвращаем номер строки. Если Элемент не найден, добавляем его в таблицу.

Public Function FindShape(shape_name As String) As Integer
Dim IShape As Shape
    FindShape = 0
    For Each IShape In ThisWorkbook.Sheets(ListN3).Shapes
            If IShape.Name = shape_name Then
                FindShape = IShape.AutoShapeType
                Exit Function
            End If
    Next IShape
End Function


Public Function FindRow(shape_name As String) As Integer
Dim row As Integer
Dim x_shape As Single
Dim y_shape As Single

    FindRow = 0
    row = 2
    With ThisWorkbook.Sheets(ListN4)
    Do While .Cells(row, 4).Value <> ""
        If .Cells(row, 4).Value = shape_name Then
            FindRow = row
            Exit Function
        End If
        row = row + 1
    Loop
'Если название элемента не найдено, добавляем его
If ThisWorkbook.Sheets(ListN3).Shapes(shape_name).AutoShapeType <> 138 Then Exit Function 'добавляем только полигоны - тип 138
.Cells(row, 1).Value = row - 1
.Cells(row, 2).Value = ThisWorkbook.Sheets(ListN3).Shapes(shape_name).ID
.Cells(row, 3).Value = ThisWorkbook.Sheets(ListN3).Shapes(shape_name).AutoShapeType
.Cells(row, 4).Value = shape_name
FindRow = row
End With
End Function

Public Function FindCol(array_name As String) As Integer
'Определяем номер столбца с набором данных по названию
Dim col As Integer
    FindCol = 0
    col = 6
    With ThisWorkbook.Sheets(ListN4)
    Do While .Cells(1, col).Value <> ""
        If .Cells(1, col).Value = array_name Then
            FindCol = col
            Exit Function
        End If
        col = col + 1
    Loop
End With
End Function

Public Sub Init()
'Добавляем названия наборов данных в "селект"
    Dim col As Integer
    Dim tmp As String
    With Worksheets(ListN3).ComboBox1
        tmp = .Value
        .Clear
        col = 6
        Do While ThisWorkbook.Sheets(ListN4).Cells(1, col).Value <> ""
            .AddItem ThisWorkbook.Sheets(ListN4).Cells(1, col).Value
            If col = 6 Then .Text = ThisWorkbook.Sheets(ListN4).Cells(1, col).Value
            col = col + 1
        Loop
    On Error GoTo linenext1
        .Value = tmp
linenext1:
    End With
End Sub

Public Sub RenewColorList(strArray As String)
'Использование палитры и легенды для схемы
Dim MinData1 As Single
Dim MaxData1 As Single
Dim tmpData1 As Single
Dim Pal As Byte
Dim Rep As Byte
Dim StepLegend As Single
Dim StepColor As Single
Dim row As Integer
Dim col As Integer
Dim tcolor As Integer
Dim LegLeft(10) As Single
Dim LegRight(10) As Single
Dim LegOn(10) As Boolean

Dim rowmax As Integer

'Ищем нужный столбец данных
col = FindCol(strArray)
If col = 0 Then Exit Sub


'Читаем легенду
With ThisWorkbook.Sheets(ListN3)
rowmax = 5 + .Cells(11, 4).Value

    For row = 6 To rowmax
            
            LegOn(row - 5) = False
            LegLeft(row - 5) = 0
            LegRight(row - 5) = 0
            
            If IsNumeric(.Cells(row, 8).Value) Then
                LegLeft(row - 5) = .Cells(row, 8).Value
                LegOn(row - 5) = True
            End If
            
            If IsNumeric(.Cells(row, 9).Value) Then
                LegRight(row - 5) = .Cells(row, 9).Value
                LegOn(row - 5) = True
            End If
            
    Next row
End With

'Показываем результат цветом в таблице данных
With ThisWorkbook.Sheets(ListN4)
row = 2
'On Error GoTo linenext1
Excel.ActiveWorkbook.Application.ScreenUpdating = False
Do While .Cells(row, 4).Value <> ""
        If .Cells(row, col).Value <> "" Then
            For Rep = 6 To rowmax
                If LegOn(Rep - 5) = True Then
                If .Cells(row, col).Value >= LegLeft(Rep - 5) Then
                If .Cells(row, col).Value <= LegRight(Rep - 5) Then
                 'MS Office 2016
                        
                        If Sheets(ListN3).Cells(Rep, 7).Interior.Pattern = xlSolid Then
                            .Cells(row, col).Interior.Pattern = xlSolid
                            .Cells(row, col).Interior.PatternColorIndex = xlAutomatic
                            .Cells(row, col).Interior.Color = Sheets(ListN3).Cells(Rep, 7).Interior.Color
                        Else
                            .Cells(row, col).Interior.Pattern = Sheets(ListN3).Cells(Rep, 7).Interior.Pattern
                            .Cells(row, col).Interior.PatternColorIndex = Sheets(ListN3).Cells(Rep, 7).Interior.PatternColorIndex
                            .Cells(row, col).Interior.ThemeColor = Sheets(ListN3).Cells(Rep, 7).Interior.ThemeColor
                            .Cells(row, col).Interior.TintAndShade = Sheets(ListN3).Cells(Rep, 7).Interior.TintAndShade
                            .Cells(row, col).Interior.PatternTintAndShade = Sheets(ListN3).Cells(Rep, 7).Interior.PatternTintAndShade
                        End If
                        
                    Exit For
                
                End If
                End If
                End If
            Next Rep
        Else
                    'MS Office 2016
                    With .Cells(row, col).Interior
                        .Pattern = xlNone
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
        End If
    row = row + 1
Loop
'linenext1:
Excel.ActiveWorkbook.Application.ScreenUpdating = True
End With
Sheets(ListN3).Select
End Sub


Public Sub CreateLegend(strArray As String)
'Создание палитры и легенды для схемы
Dim MinData1 As Single
Dim MaxData1 As Single
Dim tmpData1 As Single
Dim Pal As Byte
Dim Rep As Byte
Dim StepLegend As Single
Dim StepColor As Single
Dim row As Integer
Dim col As Integer
Dim tcolor As Integer

Pal = 5
On Error GoTo linenext1
tcolor = CInt(ThisWorkbook.Sheets(ListN3).Cells(10, 4).Interior.ThemeColor)
Pal = CByte(ThisWorkbook.Sheets(ListN3).Cells(11, 4).Value)
linenext1:
If tcolor = 0 Then
    tcolor = 4
    ThisWorkbook.Sheets(ListN3).Cells(10, 4).Interior.ThemeColor = 4
    MsgBox "Выбранный цвет не входит в тему. Установлена тема №4"
End If
col = FindCol(strArray)
If col = 0 Then Exit Sub
StepColor = 1.3 / Pal
row = 2
    
'Поиск минимума и максимума в данных
    With ThisWorkbook.Sheets(ListN4)
    .Activate
    Do While .Cells(row, 4).Value <> ""
        If .Cells(row, col).Value <> "" Then
                    tmpData1 = CSng(.Cells(row, col).Value)
            If MaxData1 < tmpData1 Then
                    MaxData1 = tmpData1
            End If
            If MinData1 > tmpData1 Then
                    MinData1 = tmpData1
            End If
        End If
        row = row + 1
    Loop

'Определение шага
StepLegend = (MaxData1 - MinData1) / Pal
row = 2
'Показываем результат цветом в таблице данных
Do While .Cells(row, 4).Value <> ""
        If .Cells(row, col).Value <> "" Then
            For Rep = 1 To Pal
                If .Cells(row, col).Value <= (MinData1 + (Rep * StepLegend)) Then
                     .Cells(row, col).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = tcolor
                            .TintAndShade = (0.8) - (StepColor * (Rep - 1))
                            .PatternTintAndShade = 0
                        End With
                    Exit For
                End If
            Next Rep
        Else
                     .Cells(row, col).Select
                    With Selection.Interior
                        .Pattern = xlNone
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
        End If
    row = row + 1
Loop
End With

Sheets(ListN3).Select
'Сначала очищаем клетки старой легенды
For Rep = 1 To 10
    Sheets(ListN3).Cells(Rep + 5, 7).Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Sheets(ListN3).Cells(Rep + 5, 8).Value = ""
    Sheets(ListN3).Cells(Rep + 5, 9).Value = ""
Next Rep

'Рисуем легенду для карты
For Rep = 1 To Pal
    Sheets(ListN3).Cells(Rep + 5, 7).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = tcolor
        .TintAndShade = (0.8) - (StepColor * (Rep - 1))
        .PatternTintAndShade = 0
    End With
    'Sheets(ListN3).Cells(Rep + 5, 8).Value = CStr(Format(MinData1 + ((Rep - 1) * StepLegend), "###0.0")) & " -- " & CStr(Format(MinData1 + (Rep * StepLegend), "###0.0"))
    Sheets(ListN3).Cells(Rep + 5, 8).Value = MinData1 + ((Rep - 1) * StepLegend)
    Sheets(ListN3).Cells(Rep + 5, 9).Value = MinData1 + (Rep * StepLegend)
Next Rep

End Sub

Public Sub ReDraw(strArray As String)
'Закрашиваем фигуры схемы соответствующими цветами
'Dim sname As String
Dim IShape As Shape
Dim row As Long
Dim Count As Long
Dim col As Integer

Init

col = FindCol(strArray)
If col = 0 Then Exit Sub
'Разгруппируем все элементы карты
    With ThisWorkbook.Sheets(ListN3)
        .Activate
        .Shapes.SelectAll
    End With
On Error GoTo linenext1
    Selection.ShapeRange.Ungroup.Select
linenext1:

ThisWorkbook.Sheets(ListN3).Shapes("LabelYear").Select
Selection.Characters.Text = strArray

    Sheets(ListN3).Activate
    
    For Each IShape In Sheets(ListN3).Shapes
        With IShape
                If .AutoShapeType = 138 Then
                    row = FindRow(.Name)
                    If row > 0 Then
                        If ThisWorkbook.Sheets(ListN4).Cells(row, col).Interior.ThemeColor > 0 Then
                            .Fill.ForeColor.ObjectThemeColor = ThisWorkbook.Sheets(ListN4).Cells(row, col).Interior.ThemeColor
                            .Fill.ForeColor.TintAndShade = ThisWorkbook.Sheets(ListN4).Cells(row, col).Interior.TintAndShade
                            .Fill.ForeColor.RGB = ThisWorkbook.Sheets(ListN4).Cells(row, col).Interior.Color
                        Else
                            'Специальная поправка для ручной легенды - простые цвета
                            .Fill.ForeColor.ObjectThemeColor = 1
                            .Fill.ForeColor.RGB = ThisWorkbook.Sheets(ListN4).Cells(row, col).Interior.Color
                        End If
                       
                    End If
                ElseIf .AutoShapeType = 1 Then
                    If Right(.Name, 3) = "lbl" Then
                        row = FindRow(Left(.Name, Len(.Name) - 4))
                        If row > 0 Then
                        'On Error GoTo linenext2
                                .DrawingObject.Characters.Text = ThisWorkbook.Sheets(ListN4).Cells(row, col + 11).Value
                                .DrawingObject.AutoSize = msoTrue
        
        
        'linenext2:
                        End If
                    End If

                End If

        End With
    Next IShape
'Группируем все элементы карты
    With ThisWorkbook.Sheets(ListN3)
        .Activate
        .Shapes.SelectAll
    End With
On Error GoTo linenext
    Selection.ShapeRange.Group.Select
    Selection.Placement = xlFreeFloating
linenext:
End Sub

Private Sub CommandButton3_Click()
Dim IShape As Shape
Dim Count As Long
Dim r As Integer
    r = MsgBox("Удалить подписи на карте?", vbOKCancel)
    If r = 1 Then
    'Разгруппируем все элементы карты
    With ThisWorkbook.Sheets(ListN3)
        .Activate
        .Shapes.SelectAll
    End With
On Error GoTo linenext1
    Selection.ShapeRange.Ungroup.Select
linenext1:
    
    For Each IShape In ActiveSheet.Shapes

        With IShape
            If .AutoShapeType = 1 Then
                If Right(.Name, 3) = "nme" Or Right(.Name, 3) = "lbl" Then
                        Count = Count + 1
                        .Delete
                End If

            End If
        End With
    Next IShape
    On Error GoTo linenext2
    Selection.ShapeRange.Group.Select
    Selection.Placement = xlFreeFloating
linenext2:
    MsgBox "Удаление " & CStr(Count) & " объектов завершено"
    End If
End Sub

Private Sub CommandButton4_Click()
'Добавляем подписи
Dim IShape As Shape
Dim IShape1 As Shape
Dim IShape2 As Shape
Dim Count As Long
Dim shape_name As String
Dim row As Integer
Dim x_shape As Single
Dim y_shape As Single
Dim col As Integer
    
col = FindCol(ComboBox1.Value)
    
ThisWorkbook.Application.ScreenUpdating = False
    'Разгруппируем все элементы карты
    With ThisWorkbook.Sheets(ListN3)
        .Activate
        .Shapes.SelectAll
    End With
On Error GoTo linenext1
    Selection.ShapeRange.Ungroup.Select
linenext1:
    
    For Each IShape In ActiveSheet.Shapes
        
            If IShape.AutoShapeType = 138 Then
                With IShape
                    shape_name = .Name
                    row = FindRow(shape_name)
                    x_shape = .Left + (.Width / 2) - 35
                    y_shape = .Top + (.Height / 2) - 22
                End With
                
                If row > 0 Then
                    
                    If FindShape(shape_name & "_lbl") = 1 Then GoTo linenext3
                    On Error GoTo linenext3
                    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, x_shape, y_shape, 40, 10).Select
                    With Selection
                                        .Name = shape_name & "_lbl"
                                        .Characters.Font.Size = 8
                                        .Characters.Text = ThisWorkbook.Sheets(ListN4).Cells(row, col + 11).Value
                                        .ShapeRange.Line.Visible = msoFalse
                                        .ShapeRange.Fill.Visible = msoFalse
                                        '.ShapeRange.Fill.Solid
                                        '.ShapeRange.Fill.ForeColor.RGB = RGB(255, 255, 255)
                                        '.ShapeRange.Fill.Transparency = 0.2
                                        .HorizontalAlignment = xlCenter
                                        .VerticalAlignment = xlCenter
                                        .AutoSize = msoTrue

                    End With
                    Set IShape1 = ActiveSheet.Shapes(shape_name & "_lbl")
                    Count = Count + 1
linenext3:

                    If FindShape(shape_name & "_nme") = 1 Then GoTo linenext5
                    On Error GoTo linenext5
                    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, x_shape, y_shape - 22, 40, 10).Select
                    With Selection
                                        .Name = shape_name & "_nme"
                                        .Characters.Font.Size = 8
                                        '.Characters.Text = shape_name
                                        If OptionButton2.Value = True Then
                                            .Characters.Text = ThisWorkbook.Sheets(ListN4).Cells(row, 5).Value
                                        Else
                                            .Characters.Text = ThisWorkbook.Sheets(ListN4).Cells(row, 4).Value
                                        End If
                                        .ShapeRange.Line.Visible = msoFalse
                                        .ShapeRange.Fill.Visible = msoFalse
                                        '.ShapeRange.Fill.Solid
                                        '.ShapeRange.Fill.ForeColor.RGB = RGB(255, 255, 255)
                                        '.ShapeRange.Fill.Transparency = 0.2
                                        .HorizontalAlignment = xlCenter
                                        .VerticalAlignment = xlCenter
                                        .AutoSize = msoTrue
                    End With
                    Set IShape2 = ActiveSheet.Shapes(shape_name & "_nme")
                    Count = Count + 1
linenext5:
                End If
            End If

    Next IShape
        With ThisWorkbook.Sheets(ListN3)
        .Activate
        .Shapes.SelectAll
    End With
    
    On Error GoTo linenext6
    Selection.ShapeRange.Group.Select
    Selection.Placement = xlFreeFloating
linenext6:
    ThisWorkbook.Application.ScreenUpdating = True
    MsgBox "Добавление " & CStr(Count) & " объектов завершено"
End Sub

Private Sub Worksheet_Activate()
    Call Init
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
