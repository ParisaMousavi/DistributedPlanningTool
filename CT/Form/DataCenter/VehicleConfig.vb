Imports Microsoft.Office.Interop.Excel

Namespace Form.DataCenter

    Friend NotInheritable Class VehicleConfig

        Public Shared ReadOnly Property VehicleHCID(Optional row As Integer = 0) As Long
            Get
                Try
                    If row = 0 Then
                        VehicleHCID = Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column).Value.ToString().Split(";")(5))
                    Else
                        VehicleHCID = Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column).Value.ToString().Split(";")(5))
                        'VehicleHCID = Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).Value.ToString().Split(";")(5))
                    End If
                Catch
                    VehicleHCID = 0
                End Try
            End Get
        End Property


        Public Shared ReadOnly Property VehiclePe57 As Long
            Get
                Try
                    VehiclePe57 = Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column).Value.ToString().Split(";")(4))
                Catch
                    VehiclePe57 = 0
                End Try

            End Get
        End Property

        Public Shared ReadOnly Property VehiclePe45(Optional row As Integer = 0) As Long
            Get
                Try
                    If row = 0 Then
                        VehiclePe45 = Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column).Value.ToString().Split(";")(3))
                    Else
                        VehiclePe45 = Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column).Value.ToString().Split(";")(3))
                    End If
                Catch
                    VehiclePe45 = 0
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property VehiclePe03(Optional row As Integer = 0) As Long
            Get
                Try
                    If row = 0 Then
                        VehiclePe03 = Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column).Value.ToString().Split(";")(2))
                    Else
                        VehiclePe03 = Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column).Value.ToString().Split(";")(2))
                    End If
                Catch
                    VehiclePe03 = 0
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property VehiclePe02(Optional row As Integer = 0) As Long
            Get
                Try
                    If row = 0 Then
                        VehiclePe02 = Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column).Value.ToString().Split(";")(1))
                    Else
                        VehiclePe02 = Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column).Value.ToString().Split(";")(1))
                    End If
                Catch
                    VehiclePe02 = 0
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property VehiclePe01 As Long
            Get
                Try
                    VehiclePe01 = Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column).Value.ToString().Split(";")(0))
                Catch
                    VehiclePe01 = 0
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property VehicleDisPlaySeq As Long
            Get
                Try
                    VehicleDisPlaySeq = Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).Value.ToString())
                Catch
                    VehicleDisPlaySeq = 0
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property VehiclePhase As String
            Get
                Try
                    VehiclePhase = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column).Value.ToString()
                Catch
                    VehiclePhase = ""
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property VehicleBuildType(Optional row As Integer = 0) As String
            Get
                Try

                    If row = 0 Then
                        VehicleBuildType = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).Value.ToString()
                    Else
                        VehicleBuildType = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).Value.ToString()
                    End If


                Catch
                    VehicleBuildType = ""
                End Try
            End Get
        End Property



        Public Shared ReadOnly Property VehicleEngine As String
            Get
                Try
                    VehicleEngine = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Engine_Column).Value.ToString()
                Catch
                    VehicleEngine = ""
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property VehicleEngineType As String
            Get
                Try
                    VehicleEngineType = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Engine_Type_Column).Value.ToString()
                Catch
                    VehicleEngineType = ""
                End Try
            End Get
        End Property


        Public Shared ReadOnly Property VehicleTransmission As String
            Get
                Try

                    VehicleTransmission = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Transmission_Column).Value.ToString()
                Catch
                    VehicleTransmission = ""
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property VehicleTransmissionType As String
            Get
                Try
                    VehicleTransmissionType = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Transmission_Type_Column).Value.ToString()
                Catch
                    VehicleTransmissionType = ""
                End Try
            End Get
        End Property



        Public Shared ReadOnly Property Rig_RigCustomerPickDate(Optional row As Integer = 0) As String
            Get
                Try
                    If row = 0 Then
                        Rig_RigCustomerPickDate = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Rig_RigCustomerPickDate_Column).Value.ToString()
                    Else
                        Rig_RigCustomerPickDate = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Rig_RigCustomerPickDate_Column).Value.ToString()
                    End If
                Catch
                    Rig_RigCustomerPickDate = ""
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Rig_CustomerRequiredDate(Optional row As Integer = 0) As String
            Get
                Try
                    If row = 0 Then
                        Rig_CustomerRequiredDate = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Rig_CustomerRequiredDate_Column).Value.ToString()
                    Else
                        Rig_CustomerRequiredDate = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Rig_CustomerRequiredDate_Column).Value.ToString()
                    End If
                Catch
                    Rig_CustomerRequiredDate = ""
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property VehicleShipToCustomer(Optional row As Integer = 0) As String
            Get
                Try
                    If row = 0 Then
                        VehicleShipToCustomer = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column).Value.ToString()
                    Else
                        VehicleShipToCustomer = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column).Value.ToString()
                    End If
                Catch
                    VehicleShipToCustomer = ""
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property VehicleSpecificationCBG(Optional row As Integer = 0) As String
            Get
                Try
                    If row = 0 Then
                        VehicleSpecificationCBG = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Specification_CBG_Column).Value.ToString()
                    Else
                        VehicleSpecificationCBG = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Specification_CBG_Column).Value.ToString()
                    End If
                Catch
                    VehicleSpecificationCBG = ""
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property VehicleXCCTeamName(Optional row As Integer = 0) As String
            Get
                Try
                    If row = 0 Then
                        VehicleXCCTeamName = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column).Value.ToString()
                    Else
                        VehicleXCCTeamName = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column).Value.ToString()
                    End If
                Catch
                    VehicleXCCTeamName = ""
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property VehicleTeamName(Optional row As Integer = 0) As String
            Get
                Try
                    If row = 0 Then
                        VehicleTeamName = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column).Value.ToString()
                    Else
                        VehicleTeamName = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column).Value.ToString()
                    End If
                Catch
                    VehicleTeamName = ""
                End Try
            End Get
        End Property


        Public Shared ReadOnly Property VehicleRemarks(Optional row As Integer = 0) As String
            Get
                Try
                    If row = 0 Then
                        VehicleRemarks = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Remarks_Column).Value.ToString()
                    Else
                        VehicleRemarks = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Remarks_Column).Value.ToString()
                    End If
                Catch
                    VehicleRemarks = ""
                End Try
            End Get
        End Property



        Public Shared ReadOnly Property VehicleDedicatedShared(Optional row As Integer = 0) As String
            Get
                Try
                    If row = 0 Then
                        VehicleDedicatedShared = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_dedicated_Shared_deleted_Column).Value.ToString()
                    Else
                        VehicleDedicatedShared = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_dedicated_Shared_deleted_Column).Value.ToString()
                    End If
                Catch
                    VehicleDedicatedShared = ""
                End Try
            End Get
        End Property


        Public Shared ReadOnly Property VehicleNumberPrefix(Optional row As Integer = 0) As String
            Get
                Try
                    If row = 0 Then
                        VehicleNumberPrefix = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Prefix_Column).Value.ToString()
                    Else
                        VehicleNumberPrefix = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Prefix_Column).Value.ToString()
                    End If
                Catch
                    VehicleNumberPrefix = ""
                End Try
            End Get
        End Property


        Public Shared ReadOnly Property VehicleBuildId(Optional row As Integer = 0) As String
            Get
                Try
                    If row = 0 Then
                        VehicleBuildId = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Build_Id_Column).Value.ToString()
                    Else
                        VehicleBuildId = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Build_Id_Column).Value.ToString()
                    End If
                Catch
                    VehicleBuildId = ""
                End Try
            End Get
        End Property


        Public Shared ReadOnly Property VehicleTagNumber(Optional row As Integer = 0) As String
            Get
                Try
                    If row = 0 Then
                        VehicleTagNumber = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Tag_Number_Column).Value.ToString()
                    Else
                        VehicleTagNumber = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Tag_Number_Column).Value.ToString()
                    End If
                Catch
                    VehicleTagNumber = ""
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property VehiclePaintFacility(Optional row As Integer = 0) As String
            Get
                Try
                    If row = 0 Then
                        VehiclePaintFacility = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Paint_Facility_Column).Value.ToString()
                    Else
                        VehiclePaintFacility = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Paint_Facility_Column).Value.ToString()
                    End If
                Catch
                    VehiclePaintFacility = ""
                End Try
            End Get
        End Property



        Public Shared ReadOnly Property VehicleNumber(Optional row As Integer = 0) As String
            Get
                Try
                    If row = 0 Then
                        VehicleNumber = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Column).Value.ToString()
                    Else
                        VehicleNumber = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Column).Value.ToString()
                    End If
                Catch
                    VehicleNumber = ""
                End Try
            End Get
        End Property


        Public Shared ReadOnly Property VehicleVin(Optional row As Integer = 0) As String
            Get
                Try
                    If row = 0 Then
                        VehicleVin = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vin_Column).Value.ToString()
                    Else
                        VehicleVin = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vin_Column).Value.ToString()
                    End If
                Catch
                    VehicleVin = ""
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property VehicleEmissionStage(Optional row As Integer = 0) As String
            Get
                Try
                    If row = 0 Then
                        VehicleEmissionStage = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Emission_Stage_Column).Value.ToString()
                    Else
                        VehicleEmissionStage = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Emission_Stage_Column).Value.ToString()
                    End If
                Catch
                    VehicleEmissionStage = ""
                End Try
            End Get
        End Property


        Public Shared ReadOnly Property VehicleBodystyle(Optional row As Integer = 0) As String
            Get
                Try
                    If row = 0 Then
                        VehicleBodystyle = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Bodystyle_Column).Value.ToString()
                    Else
                        VehicleBodystyle = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Bodystyle_Column).Value.ToString()
                    End If
                Catch
                    VehicleBodystyle = ""
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property VehicleColor(Optional row As Integer = 0) As String
            Get
                Try
                    If row = 0 Then
                        VehicleColor = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Color_Column).Value.ToString()
                    Else
                        VehicleColor = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Color_Column).Value.ToString()
                    End If
                Catch
                    VehicleColor = ""
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property VehicleDriveside(Optional row As Integer = 0) As String
            Get
                Try
                    If row = 0 Then
                        VehicleDriveside = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(Globals.ThisAddIn.Application.ActiveCell.Row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Driveside_Column).Value.ToString()
                    Else
                        VehicleDriveside = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).cells(row.ToString(), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Driveside_Column).Value.ToString()
                    End If
                Catch
                    VehicleDriveside = ""
                End Try
            End Get
        End Property



        Public Shared Function ID2Row(ID As Integer) As Long

            Dim findRange As Excel.Range = Nothing


            findRange = DirectCast(Form.DataCenter.GlobalValues.WS.Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).entirecolumn, Excel.Range).Find(ID.ToString, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False,
                Type.Missing, Type.Missing)


            ID2Row = If(findRange IsNot Nothing, findRange.Row, 0)
        End Function


        Public Shared Function Pe0345572Row(SearchText As String) As Integer

            Dim findRange As Excel.Range = Nothing

            With Form.DataCenter.GlobalValues.WS
                findRange = .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column).EntireColumn
            End With

            findRange = findRange.Find(SearchText, , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False)

            'Find(SearchText, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)

            Pe0345572Row = If(findRange IsNot Nothing, findRange.Row, 0)
        End Function
    End Class
End Namespace
