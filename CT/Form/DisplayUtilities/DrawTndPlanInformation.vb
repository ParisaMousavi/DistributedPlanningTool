
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data
Namespace Form.DisplayUtilities
    Public Class DrawTndPlanInformation
        ''' <summary>
        ''' This methode loads Program Info and 7Tabs information to the right position 
        ''' on the worksheet.
        ''' After calling this method the method ApplyFormattingAfterLoading must be called to 
        ''' apply the formating.
        ''' </summary>
        Public Function LoadTndPlanInformationToWorkSheet(UpperBoundDisplaySeq As Object, LowerBoundDisplaySeq As Object, Optional intStRow As Integer = 0, Optional intEndRow As Integer = 0) As String

            Dim StartRowInExcel As Integer
            Dim EndRowInExcel As Integer
            Dim _TndPlanInformationArray As String(,) = Nothing

            'Dim _TndPlanInformation As CT.Data.VehiclePlan.Segment.Leftside = New Data.VehiclePlan.Segment.Leftside()
            Dim _TndPlanInformationInterface As Data.Interfaces.LeftInterface

            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _TndPlanInformationInterface = New Data.VehiclePlan.Segment.Leftside
            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _TndPlanInformationInterface = New Data.RigPlan.Segment.Leftside
            Else
                Exit Function
            End If

            '----------------------------------------------------
            ' Error checking and return the error
            '----------------------------------------------------
            LoadTndPlanInformationToWorkSheet = String.Empty
            Try
                '-----------------------------------------------------------------
                ' Loading generic & specific
                '-----------------------------------------------------------------
                If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                    _TndPlanInformationArray = _TndPlanInformationInterface.GetPlanDataHcIdGeneric(DataCenter.ProgramConfig.HCID, DataCenter.ProgramConfig.BuildType, UpperBoundDisplaySeq, LowerBoundDisplaySeq)
                ElseIf Form.DataCenter.ProgramConfig.IsGeneric = False Then
                    _TndPlanInformationArray = _TndPlanInformationInterface.GetPlanDataHcIdSpecific(DataCenter.ProgramConfig.HCID, UpperBoundDisplaySeq, LowerBoundDisplaySeq, DataCenter.ProgramConfig.BuildType)
                End If

                '-------------------------------------------------
                ' error checking 
                '-------------------------------------------------
                If _TndPlanInformationArray Is Nothing Then Throw New Exception("The return value from DB in LoadTndPlanInformationToWorkSheet is empty.")


                If UpperBoundDisplaySeq IsNot Nothing AndAlso LowerBoundDisplaySeq IsNot Nothing Then
                    StartRowInExcel = intStRow
                    EndRowInExcel = intEndRow
                Else
                    StartRowInExcel = 5
                    EndRowInExcel = 5 + _TndPlanInformationArray.GetUpperBound(0)
                End If

                Dim top As Excel.Range = Form.DataCenter.GlobalValues.WS.Cells(StartRowInExcel, 2)
                Dim bottom As Excel.Range = Form.DataCenter.GlobalValues.WS.Cells(EndRowInExcel, _TndPlanInformationArray.GetUpperBound(1) + 2)
                Dim all As Excel.Range = DirectCast(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                all.Value2 = _TndPlanInformationArray

            Catch ex As Exception
                LoadTndPlanInformationToWorkSheet = ex.Message
            End Try

        End Function

        Public Function ApplyFormattingAfterLoading() As String

            Dim all As Excel.Range, top As Excel.Range, bottom As Excel.Range
            Dim rng As Excel.Range
            Dim obj As Object
            Dim strColor As String
            Dim rgbs As String()


            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = False

            ApplyFormattingAfterLoading = String.Empty
            Try


#Region "VehicleProgramInfoSection formating"
                If DataCenter.GlobalSections.VehicleProgramInfoSection IsNot Nothing Then
                    rng = Form.DataCenter.GlobalValues.WS.Cells(1, DataCenter.GlobalSections.VehicleProgramInfoSection.Cells(1, 1).Column)
                    obj = rng.Value2
                    If obj IsNot Nothing Then
                        strColor = obj.ToString()
                        rgbs = strColor.Split(";"c)
                        If rgbs.Length = 2 Then
                            Form.DataCenter.GlobalValues.WS.Columns(DataCenter.GlobalSections.VehicleProgramInfoSection.Cells(1, 1).Column).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Columns(DataCenter.GlobalSections.VehicleProgramInfoSection.Cells(1, DataCenter.GlobalSections.VehicleProgramInfoSection.Columns.Count).Column).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))
                            top = Form.DataCenter.GlobalValues.WS.Cells(5, DataCenter.GlobalSections.VehicleProgramInfoSection.Cells(1, 1).Column)
                            bottom = Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count, DataCenter.GlobalSections.VehicleProgramInfoSection.Cells(1, 1).Column)
                            all = DirectCast(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                            all.Value2 = ""
                            top = Form.DataCenter.GlobalValues.WS.Cells(5, DataCenter.GlobalSections.VehicleProgramInfoSection.Cells(1, DataCenter.GlobalSections.VehicleProgramInfoSection.Columns.Count).Column)
                            bottom = Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count, DataCenter.GlobalSections.VehicleProgramInfoSection.Cells(1, DataCenter.GlobalSections.VehicleProgramInfoSection.Columns.Count).Column)
                            all = DirectCast(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                            all.Value2 = ""
                        End If
                    End If
                End If
#End Region

#Region "InstrumentationSection formating"
                If Form.DataCenter.ProgramConfig.IsGeneric = False Then
                    rng = Form.DataCenter.GlobalValues.WS.Cells(1, DataCenter.GlobalSections.InstrumentationSection.Cells(1, 1).Column)
                    obj = rng.Value2
                    If obj IsNot Nothing Then
                        strColor = obj.ToString()
                        rgbs = strColor.Split(";"c)
                        If rgbs.Length = 2 Then
                            Form.DataCenter.GlobalValues.WS.Columns(DataCenter.GlobalSections.InstrumentationSection.Cells(1, 1).Column).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))

                            Form.DataCenter.GlobalValues.WS.Columns(DataCenter.GlobalSections.InstrumentationSection.Cells(1, DataCenter.GlobalSections.InstrumentationSection.Columns.Count).Column).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))

                            top = Form.DataCenter.GlobalValues.WS.Cells(5, DataCenter.GlobalSections.InstrumentationSection.Cells(1, 1).Column)
                            bottom = Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count, DataCenter.GlobalSections.InstrumentationSection.Cells(1, 1).Column)
                            all = DirectCast(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                            all.Value2 = ""

                            top = Form.DataCenter.GlobalValues.WS.Cells(5, DataCenter.GlobalSections.InstrumentationSection.Cells(1, DataCenter.GlobalSections.InstrumentationSection.Columns.Count).Column)
                            bottom = Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count, DataCenter.GlobalSections.InstrumentationSection.Cells(1, DataCenter.GlobalSections.InstrumentationSection.Columns.Count).Column)
                            all = DirectCast(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                            all.Value2 = ""
                        End If
                    End If
                End If
#End Region

#Region "NonMfSpecificationSection formating"
                If Form.DataCenter.ProgramConfig.IsGeneric = False Then
                    rng = Form.DataCenter.GlobalValues.WS.Cells(1, DataCenter.GlobalSections.NonMfcSpecificationSection.Cells(1, 1).Column)
                    obj = rng.Value2
                    If obj IsNot Nothing Then
                        strColor = obj.ToString()
                        rgbs = strColor.Split(";"c)
                        If rgbs.Length = 2 Then
                            Form.DataCenter.GlobalValues.WS.Columns(DataCenter.GlobalSections.NonMfcSpecificationSection.Cells(1, 1).Column).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Columns(DataCenter.GlobalSections.NonMfcSpecificationSection.Cells(1, DataCenter.GlobalSections.NonMfcSpecificationSection.Columns.Count).Column).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))
                            top = Form.DataCenter.GlobalValues.WS.Cells(5, DataCenter.GlobalSections.NonMfcSpecificationSection.Cells(1, 1).Column)
                            bottom = Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count, DataCenter.GlobalSections.NonMfcSpecificationSection.Cells(1, 1).Column)
                            all = DirectCast(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                            all.Value2 = ""
                            top = Form.DataCenter.GlobalValues.WS.Cells(5, DataCenter.GlobalSections.NonMfcSpecificationSection.Cells(1, DataCenter.GlobalSections.NonMfcSpecificationSection.Columns.Count).Column)
                            bottom = Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count, DataCenter.GlobalSections.NonMfcSpecificationSection.Cells(1, DataCenter.GlobalSections.NonMfcSpecificationSection.Columns.Count).Column)
                            all = DirectCast(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                            all.Value2 = ""
                        End If
                    End If
                End If
#End Region

#Region "MfcSpecificationSection formating"
                If Form.DataCenter.ProgramConfig.IsGeneric = False Then
                    rng = Form.DataCenter.GlobalValues.WS.Cells(1, DataCenter.GlobalSections.MfcSpecificationSection.Cells(1, 1).Column)
                    obj = rng.Value2
                    If obj IsNot Nothing Then
                        strColor = obj.ToString()
                        rgbs = strColor.Split(";"c)
                        If rgbs.Length = 2 Then
                            Form.DataCenter.GlobalValues.WS.Columns(DataCenter.GlobalSections.MfcSpecificationSection.Cells(1, 1).Column).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Columns(DataCenter.GlobalSections.MfcSpecificationSection.Cells(1, DataCenter.GlobalSections.MfcSpecificationSection.Columns.Count).Column).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))
                            top = Form.DataCenter.GlobalValues.WS.Cells(5, DataCenter.GlobalSections.MfcSpecificationSection.Cells(1, 1).Column)
                            bottom = Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count, DataCenter.GlobalSections.MfcSpecificationSection.Cells(1, 1).Column)
                            all = DirectCast(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                            all.Value2 = ""
                            top = Form.DataCenter.GlobalValues.WS.Cells(5, DataCenter.GlobalSections.MfcSpecificationSection.Cells(1, DataCenter.GlobalSections.MfcSpecificationSection.Columns.Count).Column)
                            bottom = Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count, DataCenter.GlobalSections.MfcSpecificationSection.Cells(1, DataCenter.GlobalSections.MfcSpecificationSection.Columns.Count).Column)
                            all = DirectCast(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                            all.Value2 = ""
                        End If
                    End If
                End If
#End Region

#Region "ProgramInformationSection formating"
                If Form.DataCenter.ProgramConfig.IsGeneric = False Then
                    rng = Form.DataCenter.GlobalValues.WS.Cells(1, DataCenter.GlobalSections.ProgramInformationSection.Cells(1, 1).Column)
                    obj = rng.Value2
                    If obj IsNot Nothing Then
                        strColor = obj.ToString()
                        rgbs = strColor.Split(";"c)
                        If rgbs.Length = 2 Then
                            Form.DataCenter.GlobalValues.WS.Columns(DataCenter.GlobalSections.ProgramInformationSection.Cells(1, 1).Column).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Columns(DataCenter.GlobalSections.ProgramInformationSection.Cells(1, DataCenter.GlobalSections.ProgramInformationSection.Columns.Count).Column).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))
                            top = Form.DataCenter.GlobalValues.WS.Cells(5, DataCenter.GlobalSections.ProgramInformationSection.Cells(1, 1).Column)
                            bottom = Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count, DataCenter.GlobalSections.ProgramInformationSection.Cells(1, 1).Column)
                            all = DirectCast(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                            all.Value2 = ""
                            top = Form.DataCenter.GlobalValues.WS.Cells(5, DataCenter.GlobalSections.ProgramInformationSection.Cells(1, DataCenter.GlobalSections.ProgramInformationSection.Columns.Count).Column)
                            bottom = Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count, DataCenter.GlobalSections.ProgramInformationSection.Cells(1, DataCenter.GlobalSections.ProgramInformationSection.Columns.Count).Column)
                            all = DirectCast(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                            all.Value2 = ""
                        End If
                    End If
                End If
#End Region

#Region "FurtherBasicInformationSection formating"
                If Form.DataCenter.ProgramConfig.IsGeneric = False Then
                    rng = Form.DataCenter.GlobalValues.WS.Cells(1, DataCenter.GlobalSections.FurtherBasicInformationSection.Cells(1, 1).Column)
                    obj = rng.Value2
                    If obj IsNot Nothing Then
                        strColor = obj.ToString()
                        rgbs = strColor.Split(";"c)
                        If rgbs.Length = 2 Then
                            Form.DataCenter.GlobalValues.WS.Columns(DataCenter.GlobalSections.FurtherBasicInformationSection.Cells(1, 1).Column).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Columns(DataCenter.GlobalSections.FurtherBasicInformationSection.Cells(1, DataCenter.GlobalSections.FurtherBasicInformationSection.Columns.Count).Column).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))
                            top = Form.DataCenter.GlobalValues.WS.Cells(5, DataCenter.GlobalSections.FurtherBasicInformationSection.Cells(1, 1).Column)
                            bottom = Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count, DataCenter.GlobalSections.FurtherBasicInformationSection.Cells(1, 1).Column)
                            all = DirectCast(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                            all.Value2 = ""
                            top = Form.DataCenter.GlobalValues.WS.Cells(5, DataCenter.GlobalSections.FurtherBasicInformationSection.Cells(1, DataCenter.GlobalSections.FurtherBasicInformationSection.Columns.Count).Column)
                            bottom = Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count, DataCenter.GlobalSections.FurtherBasicInformationSection.Cells(1, DataCenter.GlobalSections.FurtherBasicInformationSection.Columns.Count).Column)
                            all = DirectCast(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                            all.Value2 = ""
                        End If
                    End If
                End If
#End Region

#Region "UserShippingDetailsSection formating"
                If Form.DataCenter.ProgramConfig.IsGeneric = False Then
                    rng = Form.DataCenter.GlobalValues.WS.Cells(1, DataCenter.GlobalSections.UserShippingDetailsSection.Cells(1, 1).Column)
                    obj = rng.Value2
                    If obj IsNot Nothing Then
                        strColor = obj.ToString()
                        rgbs = strColor.Split(";"c)
                        If rgbs.Length = 2 Then
                            Form.DataCenter.GlobalValues.WS.Columns(DataCenter.GlobalSections.UserShippingDetailsSection.Cells(1, 1).Column).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Columns(DataCenter.GlobalSections.UserShippingDetailsSection.Cells(1, DataCenter.GlobalSections.UserShippingDetailsSection.Columns.Count).Column).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))
                            top = Form.DataCenter.GlobalValues.WS.Cells(5, DataCenter.GlobalSections.UserShippingDetailsSection.Cells(1, 1).Column)
                            bottom = Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count, DataCenter.GlobalSections.UserShippingDetailsSection.Cells(1, 1).Column)
                            all = DirectCast(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                            all.Value2 = ""
                            top = Form.DataCenter.GlobalValues.WS.Cells(5, DataCenter.GlobalSections.UserShippingDetailsSection.Cells(1, DataCenter.GlobalSections.UserShippingDetailsSection.Columns.Count).Column)
                            bottom = Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count, DataCenter.GlobalSections.UserShippingDetailsSection.Cells(1, DataCenter.GlobalSections.UserShippingDetailsSection.Columns.Count).Column)
                            all = DirectCast(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                            all.Value2 = ""
                        End If
                    End If
                End If
#End Region

#Region "UpdatePackSection formating"
                If Form.DataCenter.ProgramConfig.IsGeneric = False Then
                    rng = Form.DataCenter.GlobalValues.WS.Cells(1, DataCenter.GlobalSections.UpdatePackSection.Cells(1, 1).Column)
                    obj = rng.Value2
                    If obj IsNot Nothing Then
                        strColor = obj.ToString()
                        rgbs = strColor.Split(";"c)
                        If rgbs.Length = 2 Then
                            Form.DataCenter.GlobalValues.WS.Columns(DataCenter.GlobalSections.UpdatePackSection.Cells(1, 1).Column).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Columns(DataCenter.GlobalSections.UpdatePackSection.Cells(1, DataCenter.GlobalSections.UpdatePackSection.Columns.Count).Column).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))
                            top = Form.DataCenter.GlobalValues.WS.Cells(5, DataCenter.GlobalSections.UpdatePackSection.Cells(1, 1).Column)
                            bottom = Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count, DataCenter.GlobalSections.UpdatePackSection.Cells(1, 1).Column)
                            all = DirectCast(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                            all.Value2 = ""
                            top = Form.DataCenter.GlobalValues.WS.Cells(5, DataCenter.GlobalSections.UpdatePackSection.Cells(1, DataCenter.GlobalSections.UpdatePackSection.Columns.Count).Column)
                            bottom = Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count, DataCenter.GlobalSections.UpdatePackSection.Cells(1, DataCenter.GlobalSections.UpdatePackSection.Columns.Count).Column)
                            all = DirectCast(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                            all.Value2 = ""
                        End If
                    End If
                End If
#End Region

            Catch ex As Exception
                ApplyFormattingAfterLoading = "ApplyFormattingAfterLoading: " + ex.Message
            End Try


        End Function

    End Class
End Namespace
