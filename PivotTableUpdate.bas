
'       PivotTable Auto-formatting module by David Leyden
'
'
'
Const BorderColor As Long = XlRgbColor.RgbLightGray
Const TableColor1 As Long = XlRgbColor.rgbAliceBlue
Const TableColor2 As Long = XlRgbColor.rgbAntiqueWhite

Const HeadingFontSize As Integer = 18
Const TableFontSize As Integer = 14
Const ColumnOrientation As Integer = 90
Const TableRowHeight As Integer = 25

Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
    If Target.PivotCache.RecordCount > 0 Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        With Target
            .ManualUpdate = True

            .TableRange1.Interior.Color = TableColor1
            .DataBodyRange.Font.Size = TableFontSize
            .DataBodyRange.Borders(xlInsideVertical).LineStyle = 1
            .DataBodyRange.Borders(xlInsideVertical).Color = BorderColor
            .DataBodyRange.RowHeight = TableRowHeight
            .RowRange.Font.Size = 10
            .RowRange.HorizontalAlignment = -4108
            .RowRange.VerticalAlignment = -4108

            On Error Resume Next
            If .DataFields.Count > 1 Then
                With .ColumnRange
                    .Orientation = 0
                    .Font.Size = HeadingFontSize
                    .HorizontalAlignment = xlLeft
                End With
                With .DataLabelRange
                    .Orientation = ColumnOrientation
                    .Font.Size = HeadingFontSize
                    .VerticalAlignment = xlBottom
                    .HorizontalAlignment = xlCenter
                    .Borders(xlInsideVertical).Color = BorderColor
                    .Columns.AutoFit
                    .Rows.AutoFit
                End With
                .RowRange.Columns.AutoFit
                .DataBodyRange.Columns.AutoFit
                .ColumnRange.Borders(xlInsideVertical).LineStyle = xlNone
                .DataLabelRange.Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
                .DataLabelRange.Borders(xlInsideVertical).Color = BorderColor
            Else '.DataFields.Count > 1 Then
                With .ColumnRange
                    .Orientation = ColumnOrientation
                    .Font.Size = HeadingFontSize
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlBottom
                    .HorizontalAlignment = xlCenter
                    .Borders(xlInsideVertical).Color = BorderColor
                    .Columns.AutoFit
                    .Rows.AutoFit
                End With
                With .DataLabelRange
                    .Orientation = 0
                    .Font.Size = HeadingFontSize
                End With
                .DataBodyRange.Columns.AutoFit
                .RowRange.Columns.AutoFit
                .DataLabelRange.Borders(xlInsideVertical).LineStyle = xlNone
            End If
            Dim pi As PivotItem
            Dim alternator As Integer
            If .ColumnFields.Count > 1 Then
                For Each pi In .ColumnFields(1).VisibleItems
                    If Not alternator Then
                        pi.LabelRange.Interior.Color = TableColor1
                        pi.DataRange.Interior.Color = TableColor1
                        pi.ColumnRange.Interior.Color = TableColor1
                    Else
                        pi.LabelRange.Interior.Color = TableColor2
                        pi.DataRange.Interior.Color = TableColor2
                        pi.ColumnRange.Interior.Color = TableColor2
                    End If
                    alternator = Not alternator
                Next
                .ColumnRange.Borders(xlInsideVertical).LineStyle = xlNone
                .DataLabelRange.Borders(xlInsideVertical).Color = BorderColor
            Else
                .DataBodyRange.Interior.Color = TableColor1
                .ColumnRange.Interior.Color = TableColor1
                .DataLabelRange.Borders(xlInsideHorizontal).LineStyle = xlNone
                .ColumnRange.Borders(xlInsideHorizontal).LineStyle = xlNone
            End If
        End With
        Application.EnableEvents = True
        Target.ManualUpdate = False
        Application.ScreenUpdating = True
        Me.Activate
        ThisWorkbook.Windows(1).FreezePanes = False

        If Target.ColumnFields.Count > 1 Then
            Target.DataBodyRange.Select
        Else
            Me.Cells(Target.DataBodyRange.Row, 1).Select
        End If

        ThisWorkbook.Windows(1).FreezePanes = True
        Me.Range("A1").Select
    End If
End Sub



