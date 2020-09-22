Imports Microsoft.Office.Interop.Excel
Imports System.Drawing

Namespace Form.TndContextMenu
    ''' <summary>
    ''' Each Button of the context menu has a module. The Module should have at least
    ''' one public Click sub
    ''' </summary>
    Friend NotInheritable Class SelectUsercaseButton

        Public Shared Sub Click(strAddress As String)
            Try

                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.ScreenUpdating = False
                Form.DataCenter.GlobalValues.bolSelAll = False
                Form.DataCenter.GlobalValues.strSelAllAddress = ""
                Dim rngSelcolor As Excel.Range = Nothing
                Dim intRow As Integer = 0, intFCol As Integer = 0, intLCol As Integer = 0
                Dim intCnt As Integer = 0
                Dim ColSelection As New Collection
                Dim intUCSeq As Integer = 0

                With Form.DataCenter.GlobalValues.WS
                    intUCSeq = Val(.Range(strAddress).Cells(1).Formula.ToString.Split(";")(1))
                    intRow = .Range(strAddress).Row
                    Form.DisplayUtilities.Utilities.FindFLCols(intRow, intFCol, intLCol)
                    Dim rng1 As Range = Nothing, rng2 As Range = Nothing
                    intCnt = .Range(strAddress).Cells(1).column
                    Do Until intCnt = intLCol + 1
                        Try
                            If Convert.ToInt16(.Cells(intRow, intCnt).Formula.ToString.Split(";")(1)) = intUCSeq Then
                                rng1 = .Cells(intRow, intCnt)
                            End If
                        Catch ex As Exception
                        End Try
                        intCnt = intCnt + 1
                    Loop
                    rng1 = .Range(.Cells(intRow, rng1.Column), .Cells(intRow, intLCol + 1)).Find("*", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                    rng1 = .Cells(intRow, rng1.Column - 1)
                    intCnt = .Range(strAddress).Cells(1).column
                    Do Until intCnt = intFCol - 1
                        Try
                            If Convert.ToInt16(.Cells(intRow, intCnt).Formula.ToString.Split(";")(1)) = intUCSeq Then
                                rng2 = .Cells(intRow, intCnt)
                            End If
                        Catch ex As Exception
                        End Try
                        intCnt = intCnt - 1
                    Loop
                    If rng1 IsNot Nothing And rng2 IsNot Nothing Then rngSelcolor = .Range(rng1, rng2)
                    If Not rngSelcolor Is Nothing Then
                        rngSelcolor.Select()
                        If Not Form.DataCenter.ProgramConfig.ISSearchActive Then
                            Dim objFC As FormatCondition
                            With .Application.Selection
                                objFC = .FormatConditions.Add(XlFormatConditionType.xlExpression,, "=True")
                                objFC.Interior.Color = Color.White
                                objFC.Font.Color = Color.Black
                                objFC.Font.Bold = True
                            End With
                        End If
                        Form.DataCenter.GlobalValues.strUserCaseSelected = rngSelcolor.Address
                        Form.DataCenter.GlobalValues.bolUserCaseSelected = True
                    Else
                        Form.DataCenter.GlobalValues.bolUserCaseSelected = False
                        Form.DataCenter.GlobalValues.strUserCaseSelected = ""
                    End If
                    Globals.ThisAddIn.Application.EnableEvents = True
                    Globals.ThisAddIn.Application.ScreenUpdating = True
                End With
            Catch ex As Exception
                Globals.ThisAddIn.Application.EnableEvents = True
                Globals.ThisAddIn.Application.ScreenUpdating = True
            End Try
        End Sub

    End Class
End Namespace
