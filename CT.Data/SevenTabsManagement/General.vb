
Imports System.Data
Imports System.Data.SqlClient


Namespace SevenTabsManagement


    ''' <summary>
    ''' For all the stored procedures which are used for seven tabs in general
    ''' </summary>
    Public Class General
        Inherits CtBaseClass




        ''' <summary>
        ''' This function returns all the dynamic columns which belong to this HCID.
        ''' </summary>
        ''' <param name="pe01"></param>
        ''' <param name="HCID"></param>
        ''' <returns>GroupID, Header, Section, Description</returns>
        Public Function GetDynamicHeaders(pe01 As Long, HCID As Integer, MainBuildType As String) As DataTable

            Try


                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicHeaders.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe01_TnDBasicProgram_PK", SqlDbType.BigInt, 8).Value = pe01
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID


                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using


                'It means no error has been occured
                DataCenter.GlobalValues.message = String.Empty
                GetDynamicHeaders = _tbAnswer
            Catch ex As Exception
                '----------------------------------------------------------------
                ' Error classification mechanism
                '----------------------------------------------------------------
                Dim ErrorId As Integer
                Select Case ex.Message
                    Case ex.Message.IndexOf("Permission") >= 0
                        ErrorId = DataCenter.ErrorCenter.Permission
                    Case ex.Message.IndexOf("could not found") >= 0
                        ErrorId = DataCenter.ErrorCenter.Could_Not_Find_Sp
                    Case Else
                        ErrorId = DataCenter.ErrorCenter.General
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetDynamicHeaders = Nothing
            End Try


        End Function


    End Class
End Namespace