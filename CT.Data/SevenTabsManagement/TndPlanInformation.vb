
Imports System.Data
Imports System.Data.SqlClient

Namespace Reports
    Public Class TndPlanInformation

        Private _tbAnswer As DataTable = Nothing
        Private _arrayDT As String(,) = Nothing




        Public Function GetPlanData(Pe02 As Long, UpperBoundDisplaySeq As Object, LowerBoundDisplaySeq As Object) As String(,)

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Resources.ConnectionString)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedure.A3_TimeLineData_Dev.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@pe02_TnDprogramDetails_FK", SqlDbType.Int, 4).Value = Pe02
                command.Parameters.Add("@UpperBoundDisplaySeq", SqlDbType.Int, 4).Value = UpperBoundDisplaySeq
                command.Parameters.Add("@LowerBoundDisplaySeq", SqlDbType.Int, 4).Value = LowerBoundDisplaySeq


                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)
                End Using

            End Using

            ConvertDataTableToStingArray()

            GetPlanData = _arrayDT

        End Function


        ''' <summary>
        ''' The method converts the DataTable which is returned from Database to string Array 
        ''' because only array string can be written in a Excel Range.
        ''' </summary>
        Private Sub ConvertDataTableToStingArray()

            Dim i, j As Integer
            If _tbAnswer IsNot Nothing Then

                ReDim _arrayDT(_tbAnswer.Rows.Count, _tbAnswer.Columns.Count)
                For i = 0 To _tbAnswer.Rows.Count - 1
                    For j = 0 To _tbAnswer.Columns.Count - 1
                        _arrayDT(i, j) = _tbAnswer.Rows(i)(j).ToString()
                    Next j
                Next i
            End If
        End Sub


    End Class
End Namespace
