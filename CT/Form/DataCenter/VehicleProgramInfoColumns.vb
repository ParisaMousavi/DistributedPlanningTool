
Namespace Form.DataCenter
    Friend NotInheritable Class VehicleProgramInfoColumns
        Private Const col_Phase As String = "Phase"
        Private Const col_ID As String = "ID"
        Private Const col_Specification_CBG As String = "Specification CBG"
        Private Const col_HC_ID As String = "HC-ID"
        Private Const col_XCC_Team As String = "XCC Team"
        Private Const col_dedicated_shared_deleted As String = "dedicated / shared /deleted"
        Private Const col_Hardwaretype As String = "Hardwaretype"
        Private Const col_Vehicle_Number_Prefix As String = "Vehicle Number Prefix"
        Private Const col_Vehicle_Number As String = "Vehicle Number"
        Private Const col_Build_Id As String = "Build Id"
        Private Const col_Tag_Number As String = "Tag Number"
        Private Const col_Vin As String = "Vin"
        Private Const col_Engine As String = "Engine"
        Private Const col_Transmission As String = "Transmission"
        Private Const col_Emission_Stage As String = "Emission Stage"
        Private Const col_Engine_Type As String = "Engine Type"
        Private Const col_Transmission_Type As String = "Transmission Type"
        Private Const col_Bodystyle As String = "Bodystyle"
        Private Const col_Color As String = "Color"
        Private Const col_Paint_Facility As String = "Paint Facility"
        Private Const col_Driveside As String = "Driveside"
        Private Const col_Team_Names As String = "Team Names"
        Private Const col_Remarks As String = "Remarks"
        Private Const col_Ship_to_Customer As String = "Ship to Customer"
        Private Const col_P_0 As String = "P;0"
        Private Const col_CustomerRequiredDate As String = "CustomerRequiredDate"
        Private Const col_RigCustomerPickDate As String = "RigCustomerPickDate"

        Private Shared Function FindVehicleInfoColumn(strColName As String) As Integer
            Try
                Dim Rng As Excel.Range
                Dim col As Integer = 0
                Dim FCol As Integer, LCol As Integer


                FCol = Form.DataCenter.GlobalSections.VehicleProgramInfoSectionFirstColumn
                LCol = Form.DataCenter.GlobalSections.VehicleProgramInfoSectionLastColumn

                If FCol = 0 Or LCol = 0 Then
                    FindVehicleInfoFirstLastColumns(FCol, LCol)
                End If

                With Form.DataCenter.GlobalValues.WS


                    Rng = .Range(.Cells(4, FCol), .Cells(4, LCol)) 'Form.DataCenter.GlobalValues.WS.Range("4:4")


                    For Each r As Excel.Range In Rng

                        If r.Value2 = strColName Then
                            col = r.Column
                            Exit For
                        End If

                    Next


                    'Rng = .Range(.Cells(4, FCol), .Cells(4, LCol)).Find(strColName, Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, False, False)
                    '.Range(.Cells(4, FCol), .Cells(4, LCol)).Find(strColName, Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext)
                    ' .Range(.Cells(4, FCol), .Cells(4, LCol)).Find(strColName, Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext)


                End With
                'If Not Rng Is Nothing Then
                '    Return Rng.Column
                'Else
                '    Return 0
                'End If
                Return col
            Catch ex As Exception
                Return 0
            End Try
        End Function
        Public Shared Sub FindVehicleInfoFirstLastColumns(ByRef FirstCol As Integer, ByRef LastCol As Integer)
            Try
                Dim Rng1 As Excel.Range, Rng2 As Excel.Range
                With Form.DataCenter.GlobalValues.WS
                    Rng1 = .Range("2:2").Find("Program Start", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                    Rng2 = .Range("2:2").Find("Program End", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                End With
                If Rng1 IsNot Nothing And Rng2 IsNot Nothing Then
                    FirstCol = Rng1.Column
                    LastCol = Rng2.Column
                End If
            Catch ex As Exception
                FirstCol = 0
                LastCol = 0
            End Try
        End Sub
        Public Shared ReadOnly Property Vehicle_Phase_Column As Integer
            Get
                Try
                    Vehicle_Phase_Column = FindVehicleInfoColumn(col_Phase)
                Catch ex As Exception
                    Vehicle_Phase_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_ID_Column As Integer
            Get
                Try
                    Vehicle_ID_Column = FindVehicleInfoColumn(col_ID)
                Catch ex As Exception
                    Vehicle_ID_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_Specification_CBG_Column As Integer
            Get
                Try
                    Vehicle_Specification_CBG_Column = FindVehicleInfoColumn(col_Specification_CBG)
                Catch ex As Exception
                    Vehicle_Specification_CBG_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_HC_ID_Column As Integer
            Get
                Try
                    Vehicle_HC_ID_Column = FindVehicleInfoColumn(col_HC_ID)
                Catch ex As Exception
                    Vehicle_HC_ID_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_XCC_Team_Column As Integer
            Get
                Try
                    Vehicle_XCC_Team_Column = FindVehicleInfoColumn(col_XCC_Team)
                Catch ex As Exception
                    Vehicle_XCC_Team_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_dedicated_Shared_deleted_Column As Integer
            Get
                Try
                    Vehicle_dedicated_Shared_deleted_Column = FindVehicleInfoColumn(col_dedicated_shared_deleted)
                Catch ex As Exception
                    Vehicle_dedicated_Shared_deleted_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_Hardwaretype_Column As Integer
            Get
                Try
                    Vehicle_Hardwaretype_Column = FindVehicleInfoColumn(col_Hardwaretype)
                Catch ex As Exception
                    Vehicle_Hardwaretype_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_Vehicle_Number_Prefix_Column As Integer
            Get
                Try
                    Vehicle_Vehicle_Number_Prefix_Column = FindVehicleInfoColumn(col_Vehicle_Number_Prefix)
                Catch ex As Exception
                    Vehicle_Vehicle_Number_Prefix_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_Vehicle_Number_Column As Integer
            Get
                Try
                    Vehicle_Vehicle_Number_Column = FindVehicleInfoColumn(col_Vehicle_Number)
                Catch ex As Exception
                    Vehicle_Vehicle_Number_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_Build_Id_Column As Integer
            Get
                Try
                    Vehicle_Build_Id_Column = FindVehicleInfoColumn(col_Build_Id)
                Catch ex As Exception
                    Vehicle_Build_Id_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_Tag_Number_Column As Integer
            Get
                Try
                    Vehicle_Tag_Number_Column = FindVehicleInfoColumn(col_Tag_Number)
                Catch ex As Exception
                    Vehicle_Tag_Number_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_Vin_Column As Integer
            Get
                Try
                    Vehicle_Vin_Column = FindVehicleInfoColumn(col_Vin)
                Catch ex As Exception
                    Vehicle_Vin_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_Engine_Column As Integer
            Get
                Try
                    Vehicle_Engine_Column = FindVehicleInfoColumn(col_Engine)
                Catch ex As Exception
                    Vehicle_Engine_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_Transmission_Column As Integer
            Get
                Try
                    Vehicle_Transmission_Column = FindVehicleInfoColumn(col_Transmission)
                Catch ex As Exception
                    Vehicle_Transmission_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_Emission_Stage_Column As Integer
            Get
                Try
                    Vehicle_Emission_Stage_Column = FindVehicleInfoColumn(col_Emission_Stage)
                Catch ex As Exception
                    Vehicle_Emission_Stage_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_Engine_Type_Column As Integer
            Get
                Try
                    'Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = False
                    'Form.DataCenter.GlobalValues.WS.Outline.ShowLevels(0, 2) ' To find the hidden column 
                    Vehicle_Engine_Type_Column = FindVehicleInfoColumn(col_Engine_Type)
                Catch ex As Exception
                    Vehicle_Engine_Type_Column = 0
                Finally
                    'Form.DataCenter.GlobalValues.WS.Outline.ShowLevels(0, 1)
                    'Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = True
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_Transmission_Type_Column As Integer
            Get
                Try
                    Vehicle_Transmission_Type_Column = FindVehicleInfoColumn(col_Transmission_Type)
                Catch ex As Exception
                    Vehicle_Transmission_Type_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_Bodystyle_Column As Integer
            Get
                Try
                    Vehicle_Bodystyle_Column = FindVehicleInfoColumn(col_Bodystyle)
                Catch ex As Exception
                    Vehicle_Bodystyle_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Rig_RigCustomerPickDate_Column As Integer
            Get
                Try
                    Rig_RigCustomerPickDate_Column = FindVehicleInfoColumn(col_RigCustomerPickDate)
                Catch ex As Exception
                    Rig_RigCustomerPickDate_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Rig_CustomerRequiredDate_Column As Integer
            Get
                Try
                    Rig_CustomerRequiredDate_Column = FindVehicleInfoColumn(col_CustomerRequiredDate)
                Catch ex As Exception
                    Rig_CustomerRequiredDate_Column = 0
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property Vehicle_Color_Column As Integer
            Get
                Try
                    Vehicle_Color_Column = FindVehicleInfoColumn(col_Color)
                Catch ex As Exception
                    Vehicle_Color_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_Paint_Facility_Column As Integer
            Get
                Try
                    Vehicle_Paint_Facility_Column = FindVehicleInfoColumn(col_Paint_Facility)
                Catch ex As Exception
                    Vehicle_Paint_Facility_Column = 0
                End Try
            End Get
        End Property


        Public Shared ReadOnly Property Vehicle_Driveside_Column As Integer
            Get
                Try
                    Vehicle_Driveside_Column = FindVehicleInfoColumn(col_Driveside)
                Catch ex As Exception
                    Vehicle_Driveside_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_Team_Names_Column As Integer
            Get
                Try
                    Vehicle_Team_Names_Column = FindVehicleInfoColumn(col_Team_Names)
                Catch ex As Exception
                    Vehicle_Team_Names_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_Remarks_Column As Integer
            Get
                Try
                    Vehicle_Remarks_Column = FindVehicleInfoColumn(col_Remarks)
                Catch ex As Exception
                    Vehicle_Remarks_Column = 0
                End Try
            End Get
        End Property
        Public Shared ReadOnly Property Vehicle_Ship_to_Customer_Column As Integer
            Get
                Try
                    Vehicle_Ship_to_Customer_Column = FindVehicleInfoColumn(col_Ship_to_Customer)
                Catch ex As Exception
                    Vehicle_Ship_to_Customer_Column = 0
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property Vehicle_P_0_Column As Integer
            Get
                Try
                    'hard coded this because the find not working for hidden column search when filter applied and used 
                    Vehicle_P_0_Column = 3 'FindVehicleInfoColumn(col_P_0) 
                Catch ex As Exception
                    Vehicle_P_0_Column = 0
                End Try
            End Get
        End Property

    End Class
End Namespace