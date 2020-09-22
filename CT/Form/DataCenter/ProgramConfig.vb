Imports Microsoft.Office
Imports Microsoft.Office.Interop.Excel

Namespace Form.DataCenter

    Friend NotInheritable Class ProgramConfig
        Public Shared Property pe01 As Long
            Get
                Try
                    pe01 = Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B2").Value)
                Catch ex As Exception
                    pe01 = 0
                End Try
            End Get
            Set(value As Long)
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B2").Value = value
            End Set
        End Property

        ''' <summary>
        ''' The first row from which the plan after header rows starts.
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property FirstRow As Long
            Get
                Try
                    FirstRow = 5
                Catch ex As Exception
                    FirstRow = 0
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property LastRow As Long
            Get
                Try
                    With Form.DataCenter.GlobalValues.WS
                        'Dim rng3 As Excel.Range = .Range(.Cells(Form.DataCenter.ProgramConfig.FirstRow, Form.DataCenter.Vehicle_P_0_Column), .Cells(.UsedRange.Rows.Count + 10, Form.DataCenter.Vehicle_P_0_Column)).Find("", , XlFindLookIn.xlFormulas, XlLookAt.xlWhole, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, False)
                        Dim rng3 As Excel.Range = .Range(.Cells(Form.DataCenter.ProgramConfig.FirstRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column), .Cells(.UsedRange.Rows.Count + 10, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column)).Find("", , XlFindLookIn.xlFormulas, XlLookAt.xlWhole, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, False)
                        LastRow = rng3.Row - 1
                    End With
                Catch ex As Exception
                    LastRow = 0
                End Try
            End Get
        End Property

        Public Shared Property pe02 As Long

            Get

                Try
                    pe02 = Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B3").Value)
                Catch ex As Exception
                    pe02 = 0
                End Try

            End Get
            Set(value As Long)

                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B3").Value = value

            End Set

        End Property

        Public Shared Property HCID As Long

            Get

                Try
                    HCID = Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B4").Value)
                Catch ex As Exception
                    HCID = 0
                End Try

            End Get
            Set(value As Long)

                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B4").Value = value

            End Set

        End Property

        Public Shared Property HCIDName As String

            Get

                Try
                    HCIDName = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B5").Value.ToString()
                Catch ex As Exception
                    HCIDName = String.Empty


                End Try

            End Get

            Set(value As String)

                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B5").Value = value

            End Set

        End Property


        Public Shared Property IsGeneric As Boolean

            Get
                Try
                    IsGeneric = If(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B6").Value.ToString() = "Generic", True, False)
                Catch ex As Exception
                    'We have considered True because specific plan has limited permission
                    IsGeneric = True
                End Try

            End Get

            Set(value As Boolean)
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B6").Value = If(value = True, "Generic", "Specific")
            End Set

        End Property

        Public Shared Property BuildType As String

            Get

                Try
                    BuildType = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B7").Value.ToString()
                Catch ex As Exception
                    BuildType = String.Empty
                End Try

            End Get

            Set(value As String)
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B7").Value = value
            End Set

        End Property

        Public Shared Property BuildPhase As String

            Get

                Try
                    BuildPhase = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B8").Value.ToString()
                Catch ex As Exception
                    BuildPhase = String.Empty
                End Try

            End Get

            Set(value As String)
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B8").Value = value
            End Set

        End Property

        Public Shared Property Region As String

            Get

                Try
                    Region = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B21").Value.ToString()
                Catch ex As Exception
                    Region = String.Empty
                End Try

            End Get

            Set(value As String)
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B21").Value = value
            End Set

        End Property

        Public Shared Property Carline As String

            Get

                Try
                    Carline = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B9").Value.ToString()
                Catch ex As Exception
                    Carline = String.Empty
                End Try

            End Get

            Set(value As String)
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B9").Value = value
            End Set

        End Property

        Public Shared Property Platform As String

            Get

                Try
                    Platform = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B10").Value.ToString()
                Catch ex As Exception
                    Platform = String.Empty
                End Try

            End Get

            Set(value As String)
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B10").Value = value
            End Set

        End Property


        Public Shared Property XccPe26 As Long

            Get

                Try
                    XccPe26 = Long.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B11").Value.ToString())
                Catch ex As Exception
                    XccPe26 = False
                End Try

            End Get

            Set(value As Long)
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B11").Value = value.ToString()
            End Set

        End Property


        Public Shared Property XccPe01 As Long

            Get

                Try
                    XccPe01 = Long.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B12").Value.ToString())
                Catch ex As Exception
                    XccPe01 = False
                End Try

            End Get

            Set(value As Long)
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B12").Value = value.ToString()
            End Set

        End Property

        Public Shared Property AssyBuildScale As Integer

            Get

                Try
                    AssyBuildScale = Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B13").Value.ToString())
                Catch ex As Exception
                    AssyBuildScale = False
                End Try

            End Get

            Set(value As Integer)
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B13").Value = value.ToString()
            End Set

        End Property

        Public Shared Property ISSearchActive As Boolean

            Get

                Try
                    ISSearchActive = If(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B14").Value.ToString() = "YES", True, False)
                Catch ex As Exception
                    ISSearchActive = False
                End Try

            End Get

            Set(value As Boolean)
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B14").Value = If(value = True, "YES", "NO")
            End Set

        End Property

        Public Shared Property IsWithCustomFormatting As Boolean

            Get
                Try
                    IsWithCustomFormatting = If(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B15").Value.ToString() = "True", True, False)
                Catch ex As Exception
                    'We have considered True because specific plan has limited permission
                    IsWithCustomFormatting = False
                End Try

            End Get

            Set(value As Boolean)
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B15").Value = If(value = True, "True", "False")
            End Set

        End Property

        ''' <summary>
        ''' This property or attribute is using for draft scenario.
        ''' If plan is Master for Checkedout this value must be True.
        ''' If plan is draft this value should be False.
        ''' </summary>
        ''' <returns></returns>
        Public Shared Property IsMainPlan As Boolean

            Get
                Try
                    IsMainPlan = If(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B19").Value.ToString() = "True", True, False)
                Catch ex As Exception
                    'We have considered True because specific plan has limited permission
                    IsMainPlan = False
                End Try

            End Get

            Set(value As Boolean)
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B19").Value = If(value = True, "True", "False")
            End Set

        End Property


        Public Shared Property MainPlanHCID As Integer

            Get

                Try
                    MainPlanHCID = Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B20").Value.ToString())
                Catch ex As Exception
                    MainPlanHCID = 0
                End Try

            End Get

            Set(value As Integer)
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B20").Value = value
            End Set

        End Property




        Public Shared Property FileStatus As String

            Get

                Try
                    FileStatus = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B22").Value.ToString()
                Catch ex As Exception
                    FileStatus = String.Empty
                End Try

            End Get

            Set(value As String)

                'Dim PlanStatues() As String = CType([Enum].GetValues(GetType(CT.Data.DataCenter.PlanStatus)), String())
                'If PlanStatues.Contains(value) = True Then
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B22").Value = value
                'Else
                '    Throw New Exception("The Plan Status is not valid")
                'End If
            End Set
        End Property

    End Class
End Namespace