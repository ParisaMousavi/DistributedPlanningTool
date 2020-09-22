
Imports Microsoft.Office.Interop.Excel
Imports System.Drawing

Namespace Form.TndContextMenu
    ''' <summary>
    ''' Each Button of the context menu has a module. The Module should have at least
    ''' one public Click sub
    ''' </summary>
    Friend NotInheritable Class SelectAllButton
        Public Shared Sub Click(strAddress As String)
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.ScreenUpdating = False

            Form.DataCenter.GlobalValues.strUserCaseSelected = ""
            Form.DataCenter.GlobalValues.bolUserCaseSelected = False

            Dim intFCol As Integer, intLCol As Integer

            With Form.DataCenter.GlobalValues.WS
                Form.DisplayUtilities.Utilities.FindFLCols(.Range(strAddress).Row, intFCol, intLCol)
                .Range(.Cells(.Range(strAddress).Row, intFCol), .Cells(.Range(strAddress).Row, intLCol)).Select()

                If Not Form.DataCenter.ProgramConfig.ISSearchActive Then
                    Dim objFC As FormatCondition
                    With .Application.Selection
                        objFC = .FormatConditions.Add(XlFormatConditionType.xlExpression,, "=True")
                        objFC.Interior.Color = Color.White
                        objFC.Font.Color = Color.Black
                        objFC.Font.Bold = True
                    End With
                End If
                Form.DataCenter.GlobalValues.bolSelAll = True
                Form.DataCenter.GlobalValues.strSelAllAddress = .Range(.Cells(.Range(strAddress).Row, intFCol), .Cells(.Range(strAddress).Row, intLCol)).Address
            End With

            Globals.ThisAddIn.Application.Application.EnableEvents = True
            Globals.ThisAddIn.Application.Application.ScreenUpdating = True
        End Sub

    End Class
End Namespace
