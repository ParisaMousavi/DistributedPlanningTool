Imports MYEXCEL = Microsoft.Office.Interop.Excel
Imports System
Imports System.ComponentModel
Imports System.Reflection
Imports myOffice = Microsoft.Office.Core

Namespace Form.DisplayUtilities

    ''' <summary>
    ''' All the methodes which do something in interface are here
    ''' </summary>
    Friend NotInheritable Class Utilities


        Private Shared _ErrorMessage As String
        Public Shared ReadOnly Property ErrorMbessage() As String
            Get
                Return _ErrorMessage
            End Get
        End Property


        Public Shared Function DisplayTodayMarker() As Boolean

            Dim shp As Excel.Shape, shp2 As Excel.Shape, AllShapes As MYEXCEL.Shapes
            Dim shpArr As Excel.ShapeRange = Nothing
            Dim sngX As Single = 0, sngY As Single = 0, sngY2 As Single = 0
            Dim rngFnd As Excel.Range = Nothing
            Dim _Worksheet As MYEXCEL.Worksheet

            Try

                rngFnd = Form.DataCenter.GlobalSections.TimeLineSection.Find(Date.Today.ToString("yyyy-MM-dd"),, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, False, Type.Missing)
                'rngFnd = Form.DataCenter.GlobalSections.TimeLineSection.Find(Date.Today.AddYears(-2).ToString("yyyy-MM-dd"),, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, False, Type.Missing)
                If Not rngFnd Is Nothing Then

                    sngX = rngFnd.Cells(1, 1).Left + (rngFnd.EntireColumn.Width / 2)
                    sngY = rngFnd.Cells(1, 1).Top
                    sngY2 = Form.DataCenter.GlobalValues.WS.UsedRange.Height


                    _Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(CT.Form.DataCenter.WorkSheet.TnDPlan.ToString()), MYEXCEL.Worksheet)
                    AllShapes = _Worksheet.Shapes

                    If _Worksheet Is Nothing And AllShapes IsNot Nothing Then Throw New Exception("Worksheet was not found!")


                    shp = AllShapes.AddConnector(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, sngX, sngY, sngX, sngY2)


                    With shp
                        .ShapeStyle = Microsoft.Office.Core.MsoShapeStyleIndex.msoLineStylePreset21
                        .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                        .Line.Weight = 3
                    End With

                    shp2 = AllShapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, shp.Left - 30, shp.Top, 60, 16)

                    With shp2
                        .ShapeStyle = Microsoft.Office.Core.MsoShapeStyleIndex.msoShapeStylePreset35
                        .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                        .TextFrame2.TextRange.Characters.Text = "Today"
                        With .TextFrame2.TextRange.Characters(1, 5).ParagraphFormat
                            .FirstLineIndent = 0
                            .Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignCenter
                        End With
                        With .TextFrame2.TextRange.Characters(1, 5).Font
                            .NameComplexScript = "+mn-cs"
                            .NameFarEast = "+mn-ea"
                            .Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                            .Fill.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorLight1
                            .Fill.ForeColor.TintAndShade = 0
                            .Fill.ForeColor.Brightness = 0
                            .Fill.Transparency = 0
                            .Fill.Solid()
                            .Size = 11
                            .Name = "+mn-lt"
                        End With
                        .TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle
                    End With
                    Dim Arr() As String = {shp.Name.ToString, shp2.Name.ToString}
                    With Form.DataCenter.GlobalValues.WS.Shapes.Range(Arr).Group()
                        .Name = "Todaylineshape"
                        .Locked = True
                    End With
                End If

                _ErrorMessage = String.Empty
                DisplayTodayMarker = True

            Catch ex As Exception
                _ErrorMessage = "Error in displaying 'Today' marker. : " & ex.Message
                DisplayTodayMarker = False
            End Try
        End Function



        Public Shared Function FindRange(StartText As String, EndText As String) As Excel.Range

            Dim StartCell As Excel.Range = Nothing
            Dim EndCell As Excel.Range = Nothing

            StartCell = Form.DataCenter.GlobalSections.SectionSection.Find(StartText, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole,
                Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False,
                Type.Missing, Type.Missing)

            StartCell = Form.DataCenter.GlobalValues.WS.Cells(StartCell.Row, StartCell.Column)
            EndCell = Form.DataCenter.GlobalSections.SectionSection.Find(EndText, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole,
                Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False,
                Type.Missing, Type.Missing)
            EndCell = Form.DataCenter.GlobalValues.WS.Cells(4, EndCell.Column)

            FindRange = Form.DataCenter.GlobalValues.WS.Range(StartCell, EndCell)

        End Function

        Public Shared Function GetDescription(enumValue As Object) As String
            Dim fi As FieldInfo = enumValue.GetType().GetField(enumValue.ToString())
            If fi IsNot Nothing Then
                Dim attrs As Object() = fi.GetCustomAttributes(Of DescriptionAttribute)
                If (attrs IsNot Nothing And attrs.Length > 0) Then
                    Return attrs(0).Description
                End If
            End If
            Return ""
        End Function
        Public Shared Function FindDisplaySeq(SearchText As String) As Excel.Range

            Dim findRange As Excel.Range = Nothing


            findRange = DirectCast(Form.DataCenter.GlobalValues.WS.Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).entirecolumn, Excel.Range).Find(SearchText, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False,
                Type.Missing, Type.Missing)


            FindDisplaySeq = findRange
        End Function

        Public Shared Function Pe0345572Row(SearchText As String) As Integer

            Dim findRange As Excel.Range = Nothing

            findRange = Form.DataCenter.GlobalValues.WS.Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column).entirecolumn
            findRange = findRange.Find(SearchText, , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False)



            Pe0345572Row = If(findRange IsNot Nothing, findRange.Row, 0)
        End Function

        ''' <summary>
        ''' Finding first and last column of the process step
        ''' In introw specified row.
        ''' </summary>
        ''' <param name="intRow"></param>
        ''' <param name="intFCol"></param>
        ''' <param name="intLCol"></param>
        Public Shared Sub FindFLCols(intRow As Integer, ByRef intFCol As Integer, ByRef intLCol As Integer)
            On Error Resume Next
            Dim rng1 As MYEXCEL.Range, rng2 As MYEXCEL.Range
            With Form.DataCenter.GlobalValues.WS
                rng1 = .Range(.Cells(intRow, Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn), .Cells(intRow, Form.DataCenter.GlobalSections.TimeLineSectionLastColumn)).Find("*", , MYEXCEL.XlFindLookIn.xlFormulas, MYEXCEL.XlLookAt.xlWhole, MYEXCEL.XlSearchOrder.xlByColumns, MYEXCEL.XlSearchDirection.xlNext)
                rng2 = .Range(.Cells(intRow, Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn), .Cells(intRow, Form.DataCenter.GlobalSections.TimeLineSectionLastColumn)).Find("*", , MYEXCEL.XlFindLookIn.xlFormulas, MYEXCEL.XlLookAt.xlWhole, MYEXCEL.XlSearchOrder.xlByColumns, MYEXCEL.XlSearchDirection.xlPrevious)
            End With
            If Not rng1 Is Nothing Then
                intFCol = rng1.Column
            Else
                intFCol = Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn
            End If
            If Not rng2 Is Nothing Then
                intLCol = rng2.Column - 1
            Else
                intLCol = Form.DataCenter.GlobalSections.TimeLineSectionLastColumn
            End If
        End Sub
    End Class
End Namespace
