Imports System.Data
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Namespace Form.Reports
    Public Class PrecheckF4TestReport


        Private _ErrorMessage As String
        Public ReadOnly Property ErrorMessage() As String
            Get
                Return _ErrorMessage
            End Get
        End Property



        Public Function Gen_PrecheckF4TestReport() As Boolean
            Try
                Dim objData As New CT.Data.VehiclePlan.Plan
                Dim DataT As System.Data.DataTable = Nothing
                DataT = objData.PrecheckF4Test(Form.DataCenter.ProgramConfig.HCID)
                If DataT IsNot Nothing Then
                    Dim strHeader() As String = Nothing
                    Dim intCnt As Integer
                    Dim _arrayDT(,) As String = ConvertDataTableToStingArray(DataT)
                    ReDim strHeader(0 To DataT.Columns.Count - 1)
                    For intCnt = 0 To UBound(strHeader)
                        strHeader(intCnt) = DataT.Columns(intCnt).ColumnName
                    Next

                    Dim WB As Excel.Workbook = Globals.ThisAddIn.Application.Workbooks.Add()
                    Dim WS As Excel.Worksheet = WB.Worksheets(1)

                    With WS
                        .Name = "F4Test Precheck Report"
                        .Range("A1").Resize(1, strHeader.GetUpperBound(0) + 1).Value2 = strHeader
                        .Range("A2").Resize(_arrayDT.GetUpperBound(0) + 1, _arrayDT.GetUpperBound(1) + 1).Value2 = _arrayDT
                        .ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, .UsedRange, , Excel.XlYesNoGuess.xlYes).Name = "Report"
                        .ListObjects("Report").TableStyle = "TableStyleMedium9"
                        .UsedRange.EntireColumn.AutoFit()
                        .Range("B:C").EntireColumn.Hidden = True
                        .Range("1:1").EntireRow.RowHeight = 20
                    End With
                    WB.Activate()
                End If

                _ErrorMessage = String.Empty
                Return True
            Catch ex As Exception
                _ErrorMessage = ex.Message
                Return False
            End Try
        End Function
        Private Function ConvertDataTableToStingArray(_tbAnswer As System.Data.DataTable) As String(,)
            Try
                Dim i, j As Integer
                Dim _arrayDT(,) As String = Nothing
                If _tbAnswer IsNot Nothing Then
                    ReDim _arrayDT(_tbAnswer.Rows.Count, _tbAnswer.Columns.Count)
                    For i = 0 To _tbAnswer.Rows.Count - 1
                        For j = 0 To _tbAnswer.Columns.Count - 1
                            _arrayDT(i, j) = _tbAnswer.Rows(i)(j).ToString()
                        Next j
                    Next i
                End If
                Return _arrayDT
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
    End Class
End Namespace
