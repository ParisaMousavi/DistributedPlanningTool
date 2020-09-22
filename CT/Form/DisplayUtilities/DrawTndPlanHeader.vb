
Imports CT.Data.Reports
Imports Excel = Microsoft.Office.Interop.Excel
Imports Office = Microsoft.Office.Core
Imports Microsoft.Office.Tools.Excel
Imports System.Data
Imports System.Windows.Forms

Namespace Form.DisplayUtilities
    Public Class DrawTndPlanHeader

        Event EventUpdateProgress(progressvalue As Double)

        Private Sub UpdateProgressbar(progressvalue As Double)
            RaiseEvent EventUpdateProgress(progressvalue)
        End Sub

        Private _ErrorMessage As String
        Public ReadOnly Property ErrorMessage() As String
            Get
                Return _ErrorMessage
            End Get
        End Property


        Public Function LoadTndPlanHeaderToWorkSheet(ByRef TotalColumn As Integer) As String
            '--------------------------------------------------------
            'error controling with message
            '--------------------------------------------------------
            LoadTndPlanHeaderToWorkSheet = String.Empty

            'Dim _TndPlanHeader As CT.Data.VehiclePlan.Segment.Header = New Data.VehiclePlan.Segment.Header()
            Dim _TndPlanHeaderInterface As Data.Interfaces.HeaderInterface

            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _TndPlanHeaderInterface = New Data.VehiclePlan.Segment.Header
            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _TndPlanHeaderInterface = New Data.RigPlan.Segment.Header
            Else
                Exit Function
            End If

            Dim _TndPlanHeaderArray As String(,) = Nothing

            Try

                Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                '---------------------------------------------------------------
                'Loading Specific & Generic
                '---------------------------------------------------------------
                If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                    _TndPlanHeaderArray = _TndPlanHeaderInterface.GetPlanHeaderGeneric(Form.DataCenter.ProgramConfig.HCID,
                                                                Form.DataCenter.ProgramConfig.BuildType,
                                                                Form.DataCenter.ProgramConfig.BuildPhase, Form.DataCenter.ProgramConfig.BuildType)
                    If _TndPlanHeaderArray Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                ElseIf Form.DataCenter.ProgramConfig.IsGeneric = False Then
                    _TndPlanHeaderArray = _TndPlanHeaderInterface.GetPlanHeaderSpecific(Form.DataCenter.ProgramConfig.HCID,
                                                                Form.DataCenter.ProgramConfig.BuildType,
                                                                Form.DataCenter.ProgramConfig.BuildPhase, Form.DataCenter.ProgramConfig.BuildType)
                    If _TndPlanHeaderArray Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                End If
                '-----------------------------------------------------
                'correctness control
                '-----------------------------------------------------
                If _TndPlanHeaderArray Is Nothing Then Throw New Exception("The return value from DB in LoadTndPlanHeaderToWorkSheet is empty.")

                Dim top As Excel.Range = Form.DataCenter.GlobalValues.WS.Cells(1, 1)
                Dim bottom As Excel.Range = Form.DataCenter.GlobalValues.WS.Cells(_TndPlanHeaderArray.GetUpperBound(0) + 1, _TndPlanHeaderArray.GetUpperBound(1) + 1)
                Dim all As Excel.Range = Form.DataCenter.GlobalValues.WS.Range(top, bottom)
                all.Value2 = _TndPlanHeaderArray
                Form.DataCenter.GlobalValues.WS.Cells(4, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column).AddComment("Input format 'dd-MM-yyyy (include single quote also).")
                TotalColumn = _TndPlanHeaderArray.GetUpperBound(1) + 1
            Catch ex As Exception

                LoadTndPlanHeaderToWorkSheet = ex.Message
            End Try

        End Function

        Public Function ApplyFormattingAfterLoading(TotalColumn As Integer, WithCustomFormat As Boolean) As Boolean
            Try
                ApplyColorAndMergeToHeaderSection(TotalColumn)
                ApplyHolidaysFlags()
                ApplyGatewayFlags()
                Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                If WithCustomFormat = True Then
                    If ApplyCustomFormatting() = False Then Throw New Exception(_ErrorMessage)
                End If
                UpdateProgressbar(5)

                _ErrorMessage = String.Empty
                ApplyFormattingAfterLoading = True
            Catch ex As Exception
                _ErrorMessage = ex.Message
                ApplyFormattingAfterLoading = False
            End Try
        End Function


        ''' <summary>
        ''' 
        ''' </summary>
        Private Function ApplyCustomFormatting() As Boolean
            Try

                '------------------------------------------------------------
                ' Check if Custom formatting Not existed  then generate it.
                '------------------------------------------------------------
                Dim dndplanHdr As DataTable = Nothing
                Dim _planindfor As CT.Data.PlanIndivitualFormatting = New CT.Data.PlanIndivitualFormatting()
                dndplanHdr = _planindfor.GetTndPlanHeaderSettings(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType, Environment.UserDomainName + "\" + Environment.UserName)
                If dndplanHdr Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                If Not dndplanHdr Is Nothing Then
                    If dndplanHdr.Rows.Count <= 0 Then
                        _planindfor.InitialFormat(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.FileStatus)
                        dndplanHdr = _planindfor.GetTndPlanHeaderSettings(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType, Environment.UserDomainName + "\" + Environment.UserName)
                    End If
                End If

                '------------------------------------------------------------
                ' Noe dndplanHdr has value and the settings must be applied on interface
                'Jeeva
                '------------------------------------------------------------
                '------------------------------------------------------------
                If dndplanHdr.Rows.Count > 0 Then
                    For i = 0 To dndplanHdr.Rows.Count - 1
                        For j = 0 To dndplanHdr.Columns.Count - 1
                            Dim Header = dndplanHdr.Rows(i)(j).ToString().Split(";")
                            If Header.Length >= 5 Then
                                Form.DataCenter.GlobalValues.WS.Columns(j + 1).entirecolumn.ColumnWidth = Header(Header.Length - 2)
                            End If
                        Next
                    Next
                End If

                _ErrorMessage = String.Empty
                ApplyCustomFormatting = True
            Catch ex As Exception
                _ErrorMessage = ex.Message
                ApplyCustomFormatting = False
            End Try
        End Function



        Public Sub ApplyHolidaysFlags()
            Dim _obj As New Form.DataCenter.ModuleFunction
            Dim holidays As CT.Data.PublicHoliday
            Dim holidaysdt As System.Data.DataTable
            Dim FindColumn As Excel.Range = Nothing
            Dim dblProgress As Double
            Try
                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.DisplayAlerts = False
                Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)

                holidays = New Data.PublicHoliday()
                holidaysdt = holidays.GetPlanPublicHolidaysForHeader(Form.DataCenter.ProgramConfig.BuildType, CT.Form.DataCenter.ProgramConfig.HCID)
                If holidaysdt Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)


                Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.WS.Cells(2, Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn), Form.DataCenter.GlobalValues.WS.Cells(2, Form.DataCenter.GlobalSections.TimeLineSectionLastColumn)).Interior.Color = Excel.Constants.xlNone
                Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.WS.Cells(2, Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn), Form.DataCenter.GlobalValues.WS.Cells(2, Form.DataCenter.GlobalSections.TimeLineSectionLastColumn)).Value = ""


                For Each dr As System.Data.DataRow In holidaysdt.Rows
                    Form.DataCenter.GlobalValues.WS.Parent.activate()
                    Form.DataCenter.GlobalValues.WS.Activate()
                    Globals.ThisAddIn.Application.ScreenUpdating = False
                    Globals.ThisAddIn.Application.EnableEvents = False
                    Globals.ThisAddIn.Application.DisplayAlerts = False
                    UpdateProgressbar(dblProgress)

                    If IsDate(dr("PublicHolidayStart")) Then

                        FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(Date.Parse(dr("PublicHolidayStart").ToString()).ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)

                        If FindColumn IsNot Nothing Then

                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(153), Integer.Parse(63), Integer.Parse(153)))
                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(255), Integer.Parse(255), Integer.Parse(255)))
                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column) = dr("PublicHolidayName").ToString


                            FindColumn = Nothing

                        End If
                    End If


                Next



            Catch ex As Exception
                _obj.sbProtectPlan()
                'System.Windows.Forms.MessageBox.Show(ex.Message, "Draw Holidays Flags in timeline", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.DrawTndPlanHeader, ex.Message), "Draw Holidays Flags in timeline", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
            Finally
                _obj.sbProtectPlan()

            End Try
        End Sub

        ''' <summary>
        ''' This function cleans the flags line and
        ''' draws all the date that we have in DB on interface.
        ''' </summary>
        Public Sub ApplyGatewayFlags()
            Dim _PlanInterface As Data.Interfaces.PlanInterface
            Dim _AddtionalDateInformation As Data.AddtionalDateInformation = New Data.AddtionalDateInformation()

            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _PlanInterface = New Data.VehiclePlan.Plan
            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _PlanInterface = New Data.RigPlan.Plan
            Else
                Exit Sub
            End If

            Dim dtTable As System.Data.DataTable
            Dim FindColumn As Excel.Range = Nothing
            Dim Backrgb, Fontrgb As String
            Dim _obj As New Form.DataCenter.ModuleFunction


            Try
                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.DisplayAlerts = False
                Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)

                dtTable = _AddtionalDateInformation.SelectAllDateInformation(Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.BuildType)
                If dtTable.Rows.Count < 1 Then
                    dtTable = _PlanInterface.SelectDateInformation(Form.DataCenter.ProgramConfig.pe02)
                    If dtTable.Rows.Count < 0 Then Throw New Exception("There is no row recorded.")
                End If


                Dim dblProgress As Double
                dblProgress = 5 / dtTable.Rows.Count
                For Each dr As System.Data.DataRow In dtTable.Rows
                    Form.DataCenter.GlobalValues.WS.Parent.activate()
                    Form.DataCenter.GlobalValues.WS.Activate()
                    Globals.ThisAddIn.Application.ScreenUpdating = False
                    Globals.ThisAddIn.Application.EnableEvents = False
                    Globals.ThisAddIn.Application.DisplayAlerts = False
                    UpdateProgressbar(dblProgress)
                    Backrgb = dr("DateBackRGB").ToString()
                    Fontrgb = dr("DateFontRGB").ToString()
                    If IsDate(dr("AssyMRD")) Then

                        FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(Date.Parse(dr("AssyMRD").ToString()).ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)

                        If FindColumn IsNot Nothing Then
                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(Backrgb.Substring(0, 3)), Integer.Parse(Backrgb.Substring(3, 3)), Integer.Parse(Backrgb.Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(Fontrgb.Substring(0, 3)), Integer.Parse(Fontrgb.Substring(3, 3)), Integer.Parse(Fontrgb.Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column) = "MRD - " + dr("HealthChartId").ToString

                            FindColumn = Nothing

                        End If
                    End If


                    'If Form.DataCenter.ProgramConfig.BuildPhase = CT.Data.DataCenter.BuildPhase.VP.ToString() Then


                    If IsDate(dr("Firstm1")) Then

                        FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(Date.Parse(dr("Firstm1").ToString()).ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)

                        If FindColumn IsNot Nothing Then

                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(Backrgb.Substring(0, 3)), Integer.Parse(Backrgb.Substring(3, 3)), Integer.Parse(Backrgb.Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(Fontrgb.Substring(0, 3)), Integer.Parse(Fontrgb.Substring(3, 3)), Integer.Parse(Fontrgb.Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column) = "1st M1 - " + dr("HealthChartId").ToString


                            FindColumn = Nothing

                        End If
                    End If

                    If IsDate(dr("M1DC")) Then

                        FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(Date.Parse(dr("M1DC").ToString()).ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)

                        If FindColumn IsNot Nothing Then

                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(Backrgb.Substring(0, 3)), Integer.Parse(Backrgb.Substring(3, 3)), Integer.Parse(Backrgb.Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(Fontrgb.Substring(0, 3)), Integer.Parse(Fontrgb.Substring(3, 3)), Integer.Parse(Fontrgb.Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column) = "M1DC - " + dr("HealthChartId").ToString


                            FindColumn = Nothing

                        End If
                    End If

                    'ElseIf Form.DataCenter.ProgramConfig.BuildPhase = CT.Data.DataCenter.BuildPhase.M1.ToString() Then

                    If IsDate(dr("FirstVP")) Then

                        FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(Date.Parse(dr("FirstVP").ToString()).ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)

                        If FindColumn IsNot Nothing Then

                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(Backrgb.Substring(0, 3)), Integer.Parse(Backrgb.Substring(3, 3)), Integer.Parse(Backrgb.Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(Fontrgb.Substring(0, 3)), Integer.Parse(Fontrgb.Substring(3, 3)), Integer.Parse(Fontrgb.Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column) = "1st VP - " + dr("HealthChartId").ToString


                            FindColumn = Nothing

                        End If
                    End If


                    If IsDate(dr("PEC")) Then

                        FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(Date.Parse(dr("PEC").ToString()).ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)

                        If FindColumn IsNot Nothing Then

                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(Backrgb.Substring(0, 3)), Integer.Parse(Backrgb.Substring(3, 3)), Integer.Parse(Backrgb.Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(Fontrgb.Substring(0, 3)), Integer.Parse(Fontrgb.Substring(3, 3)), Integer.Parse(Fontrgb.Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column) = "PEC - " + dr("HealthChartId").ToString


                            FindColumn = Nothing

                        End If
                    End If


                    If IsDate(dr("FEC")) Then

                        FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(Date.Parse(dr("FEC").ToString()).ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)

                        If FindColumn IsNot Nothing Then

                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(Backrgb.Substring(0, 3)), Integer.Parse(Backrgb.Substring(3, 3)), Integer.Parse(Backrgb.Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(Fontrgb.Substring(0, 3)), Integer.Parse(Fontrgb.Substring(3, 3)), Integer.Parse(Fontrgb.Substring(6, 3))))
                            Form.DataCenter.GlobalValues.WS.Cells(2, FindColumn.Column) = "FEC - " + dr("HealthChartId").ToString
                            FindColumn = Nothing

                        End If
                    End If

                Next
            Catch ex As Exception
                _obj.sbProtectPlan()
                'System.Windows.Forms.MessageBox.Show(ex.Message, "Draw Gateway Flags in timeline", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.DrawTndPlanHeader, ex.Message), "Draw Gateway Flags in timeline", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
            End Try
            _obj.sbProtectPlan()

        End Sub


        Private Sub ApplyColorAndMergeToHeaderSection(TotalColumn As Integer)

            Dim counter As Integer = 1
            Dim FisrtOfRange As Excel.Range = Nothing
            Dim EndOfRange As Excel.Range = Nothing
            Dim TopOfRange As Excel.Range = Nothing
            Dim ButtomOfRange As Excel.Range = Nothing
            Dim top As Excel.Range = Nothing
            Dim bottom As Excel.Range = Nothing
            Dim all As Excel.Range = Nothing
            Dim obj As Object, WB As Excel.Workbook
            Dim s As String
            WB = Form.DataCenter.GlobalValues.WS.Parent
            Try
                'Form.DataCenter.GlobalValues.TotalColumn              
                Dim dblProgress As Double
                Dim dbl_TotalProgress As Double = 0
                dblProgress = 5 / TotalColumn
                While (counter <= TotalColumn)
                    WB.Activate()
                    'Form.DataCenter.GlobalValues.WS.Activate()
                    Globals.ThisAddIn.Application.ScreenUpdating = False
                    dbl_TotalProgress += dblProgress
                    If dbl_TotalProgress > 2 Then 'update progress bar only for value crossing 2 
                        UpdateProgressbar(dbl_TotalProgress)
                        dbl_TotalProgress = 0
                    End If
                    obj = Form.DataCenter.GlobalValues.WS.Cells(2, counter).Value2
                    If (obj IsNot Nothing) Then

                        If (FisrtOfRange Is Nothing) Then
                            FisrtOfRange = Form.DataCenter.GlobalValues.WS.Cells(2, counter)
                        ElseIf (FisrtOfRange.Value2 = (Form.DataCenter.GlobalValues.WS.Cells(2, counter)).Value2) Then
                            EndOfRange = Form.DataCenter.GlobalValues.WS.Cells(2, counter)
                        ElseIf (FisrtOfRange IsNot Nothing And EndOfRange IsNot Nothing) Then
                            Form.DataCenter.GlobalValues.WS.Range(FisrtOfRange, EndOfRange).Merge(False)
                            FisrtOfRange = Form.DataCenter.GlobalValues.WS.Cells(2, counter)
                            EndOfRange = Nothing
                        Else
                            FisrtOfRange = Form.DataCenter.GlobalValues.WS.Cells(2, counter)
                            EndOfRange = Nothing
                        End If
                    End If

                    TopOfRange = Nothing
                    ButtomOfRange = Form.DataCenter.GlobalValues.WS.Cells(4, counter)

                    If ((Form.DataCenter.GlobalValues.WS.Cells(4, counter)).Value2 = (Form.DataCenter.GlobalValues.WS.Cells(3, counter)).Value2) Then
                        TopOfRange = Form.DataCenter.GlobalValues.WS.Cells(3, counter)
                        If ((Form.DataCenter.GlobalValues.WS.Cells(3, counter)).Value2 = (Form.DataCenter.GlobalValues.WS.Cells(2, counter)).Value2) Then
                            TopOfRange = Form.DataCenter.GlobalValues.WS.Cells(2, counter)
                        End If
                    End If

                    If (TopOfRange IsNot Nothing And ButtomOfRange IsNot Nothing) Then
                        Form.DataCenter.GlobalValues.WS.Range(ButtomOfRange, TopOfRange).Merge(False)
                        TopOfRange = Nothing
                        ButtomOfRange = Nothing
                    End If

                    Dim strColor As String = ""

                    Try
                        strColor = Form.DataCenter.GlobalValues.WS.Cells(1, counter).Value2.ToString()
                    Catch ex As Exception
                    End Try

                    Dim rgbs As String() = strColor.Split(";")
                    If (rgbs.Length = 2) Then
                        If Form.DataCenter.GlobalValues.WS.Cells(4, counter).MergeCells = True Then
                            CType(Form.DataCenter.GlobalValues.WS.Cells(4, counter), Excel.Range).MergeArea.Cells(1, 1).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))
                            CType(Form.DataCenter.GlobalValues.WS.Cells(4, counter), Excel.Range).MergeArea.Cells(1, 1).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(1).Substring(0, 3)), Integer.Parse(rgbs(1).Substring(3, 3)), Integer.Parse(rgbs(1).Substring(6, 3))))
                        End If

                        obj = Form.DataCenter.GlobalValues.WS.Cells(1, counter).Value2

                        If (obj IsNot Nothing) Then
                            s = Convert.ToString((Form.DataCenter.GlobalValues.WS.Cells(2, counter)).Value2)
                            If (Form.DataCenter.GlobalSections.SectionFlags.Contains(s) = True) Then
                                Form.DataCenter.GlobalValues.WS.Columns(counter).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))
                                top = Form.DataCenter.GlobalValues.WS.Cells(5, counter)
                                bottom = Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count, counter)
                                all = CType(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                                all.Value2 = ""
                            End If
                        End If
                    End If
                    counter = counter + 1
                End While
                UpdateProgressbar(dbl_TotalProgress) 'For balance progress update
                DataCenter.GlobalValues.WS.Range("A1").EntireRow.Hidden = True
                Form.DataCenter.GlobalSections.ColorSection.EntireRow.Hidden = True
                DataCenter.GlobalValues.WS.Range("A1").EntireColumn.Hidden = True
                DataCenter.GlobalValues.WS.Range("C1").EntireColumn.Hidden = True
                Form.DataCenter.GlobalSections.SectionSection.RowHeight = 75
                Form.DataCenter.GlobalSections.HeaderSection.RowHeight = 23
                Form.DataCenter.GlobalSections.DescriptionSection.RowHeight = 85

                '------------------------------------------------------------------
                ' it's for drawing gateways flag FEC, PEC and ets. here
                '------------------------------------------------------------------
                Form.DataCenter.GlobalSections.TimeLineSection.UnMerge()
                'Globals.ThisAddIn.Application.ScreenUpdating = False
                'Globals.ThisAddIn.Application.EnableEvents = False
                'Globals.ThisAddIn.Application.DisplayAlerts = False
            Catch ex As Exception
                'System.Windows.Forms.MessageBox.Show(ex.Message, "ApplyColorAndMergeToHeaderSection (Total column)", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.DrawTndPlanHeader, ex.Message), "ApplyColorAndMergeToHeaderSection (Total column)", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
            End Try

        End Sub



        ''' <summary>
        ''' Only for updating a section of specificatiob tables
        ''' </summary>
        ''' <param name="SpecificationSection"></param>
        Public Sub ApplyColorAndMergeToHeaderSection(SpecificationSection As Form.DataCenter.GlobalSections.SectionName)

            Dim counter As Integer = 1
            Dim FisrtOfRange As Excel.Range = Nothing
            Dim EndOfRange As Excel.Range = Nothing
            Dim TopOfRange As Excel.Range = Nothing
            Dim ButtomOfRange As Excel.Range = Nothing
            Dim top As Excel.Range = Nothing
            Dim bottom As Excel.Range = Nothing
            Dim all As Excel.Range = Nothing
            Dim obj As Object
            Dim s As String
            Dim TotalColumn As Integer
            Dim strOutput As String(,) = Nothing
            Dim strOutputData As String(,) = Nothing
            Dim objGlobal As New Form.DataCenter.ModuleFunction
            Dim CurrentHiddenStateOfSection As Boolean = True
            Try

                Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual
                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.DisplayAlerts = False

                Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)

                TotalColumn = 0
                counter = 0
                Select Case SpecificationSection
                    Case Form.DataCenter.GlobalSections.SectionName.InstrumentationSection
                        counter = Form.DataCenter.GlobalSections.InstrumentationSectionFirstColumn + 1
                        TotalColumn = counter + Form.DataCenter.GlobalSections.InstrumentationSection.Columns.Count - 3
                        CurrentHiddenStateOfSection = Form.DataCenter.GlobalSections.InstrumentationSection.Columns.Hidden
                        'Form.DataCenter.GlobalSections.InstrumentationSection.UnMerge()
                        '--------------------------------------------------------------------------
                        ' Fetching new header from data base
                        '--------------------------------------------------------------------------
                        Dim _Instrumentation As CT.Data.Interfaces.InstrumentationInterface = Nothing
                        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
                            _Instrumentation = New CT.Data.VehiclePlan.SevenTabs.Instrumentation
                        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
                            _Instrumentation = New CT.Data.RigPlan.SevenTabs.Instrumentation
                        End If

                        strOutput = _Instrumentation.GetTndPlanHeader(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.BuildPhase, Form.DataCenter.ProgramConfig.BuildType)
                        If strOutput Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        strOutputData = _Instrumentation.GetPlanData(Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Form.DataCenter.ProgramConfig.BuildType)
                        If strOutputData Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

                    Case Form.DataCenter.GlobalSections.SectionName.NonMfcSpecificationSection
                        counter = Form.DataCenter.GlobalSections.NonMfSpecificationSectionFirstColumn + 1
                        TotalColumn = counter + Form.DataCenter.GlobalSections.NonMfcSpecificationSection.Columns.Count - 3
                        CurrentHiddenStateOfSection = Form.DataCenter.GlobalSections.NonMfcSpecificationSection.Columns.Hidden
                        'Form.DataCenter.GlobalSections.NonMfcSpecificationSection.UnMerge()
                        '--------------------------------------------------------------------------
                        ' Fetching new header from data base
                        '--------------------------------------------------------------------------
                        Dim _NonMfcSpecification As CT.Data.Interfaces.NonMfcInterface = Nothing
                        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
                            _NonMfcSpecification = New CT.Data.VehiclePlan.SevenTabs.NonMfcSpecification
                        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
                            _NonMfcSpecification = New CT.Data.RigPlan.SevenTabs.NonMfcSpecification
                        End If


                        strOutput = _NonMfcSpecification.GetTndPlanHeader(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.BuildPhase, Form.DataCenter.ProgramConfig.BuildType)
                        If strOutput Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        strOutputData = _NonMfcSpecification.GetPlanData(Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Form.DataCenter.ProgramConfig.BuildType)
                        If strOutputData Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

                    Case Form.DataCenter.GlobalSections.SectionName.MfcSpecificationSection
                        counter = Form.DataCenter.GlobalSections.MfcSpecificationSectionFirstColumn + 1
                        TotalColumn = counter + Form.DataCenter.GlobalSections.MfcSpecificationSection.Columns.Count - 3
                        CurrentHiddenStateOfSection = Form.DataCenter.GlobalSections.MfcSpecificationSection.Columns.Hidden
                        'Form.DataCenter.GlobalSections.MfcSpecificationSection.UnMerge()

                        '--------------------------------------------------------------------------
                        ' Fetching new header from data base
                        '--------------------------------------------------------------------------
                        Dim _MfcSpecification As CT.Data.Interfaces.MfcInterface = Nothing
                        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                            _MfcSpecification = New CT.Data.VehiclePlan.SevenTabs.MfcSpecification
                        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                            _MfcSpecification = New CT.Data.RigPlan.SevenTabs.MfcSpecification
                        End If

                        strOutput = _MfcSpecification.GetTndPlanHeader(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.BuildPhase, Form.DataCenter.ProgramConfig.BuildType)
                        If strOutput Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        strOutputData = _MfcSpecification.GetPlanData(Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Form.DataCenter.ProgramConfig.BuildType)
                        If strOutputData Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

                    Case Form.DataCenter.GlobalSections.SectionName.ProgramInformationSection
                        counter = Form.DataCenter.GlobalSections.ProgramInformationSectionFirstColumn + 1
                        TotalColumn = counter + Form.DataCenter.GlobalSections.ProgramInformationSection.Columns.Count - 3
                        CurrentHiddenStateOfSection = Form.DataCenter.GlobalSections.ProgramInformationSection.Columns.Hidden
                        'Form.DataCenter.GlobalSections.ProgramInformationSection.UnMerge()
                        '--------------------------------------------------------------------------
                        ' Fetching new header from data base
                        '--------------------------------------------------------------------------
                        Dim _ProgramInformation As CT.Data.Interfaces.ProgramInformationInterface = Nothing
                        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
                            _ProgramInformation = New CT.Data.VehiclePlan.SevenTabs.ProgramInformation
                        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
                            _ProgramInformation = New CT.Data.RigPlan.SevenTabs.ProgramInformation
                        End If

                        strOutput = _ProgramInformation.GetTndPlanHeader(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.BuildPhase, Form.DataCenter.ProgramConfig.BuildType)
                        If strOutput Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        strOutputData = _ProgramInformation.GetPlanData(Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Form.DataCenter.ProgramConfig.BuildType)
                        If strOutputData Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

                    Case Form.DataCenter.GlobalSections.SectionName.FurtherBasicInformationSection
                        counter = Form.DataCenter.GlobalSections.FurtherBasicInformationSectionFirstColumn + 1
                        TotalColumn = counter + Form.DataCenter.GlobalSections.FurtherBasicInformationSection.Columns.Count - 3
                        CurrentHiddenStateOfSection = Form.DataCenter.GlobalSections.FurtherBasicInformationSection.Columns.Hidden
                        'Form.DataCenter.GlobalSections.FurtherBasicInformationSection.UnMerge()
                        '--------------------------------------------------------------------------
                        ' Fetching new header from data base
                        '--------------------------------------------------------------------------
                        Dim _FurtherBasicSpecification As CT.Data.Interfaces.FurtherBasicInterface = Nothing
                        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
                            _FurtherBasicSpecification = New CT.Data.VehiclePlan.SevenTabs.FurtherBasicSpecification
                        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
                            _FurtherBasicSpecification = New CT.Data.RigPlan.SevenTabs.FurtherBasicSpecification
                        End If

                        strOutput = _FurtherBasicSpecification.GetTndPlanHeader(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.BuildPhase, Form.DataCenter.ProgramConfig.BuildType)
                        If strOutput Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        strOutputData = _FurtherBasicSpecification.GetPlanData(Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Form.DataCenter.ProgramConfig.BuildType)
                        If strOutputData Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

                    Case Form.DataCenter.GlobalSections.SectionName.UserShippingDetailsSection
                        counter = Form.DataCenter.GlobalSections.UserShippingDetailsSectionFirstColumn + 1
                        TotalColumn = counter + Form.DataCenter.GlobalSections.UserShippingDetailsSection.Columns.Count - 3
                        CurrentHiddenStateOfSection = Form.DataCenter.GlobalSections.UserShippingDetailsSection.Columns.Hidden
                        'Form.DataCenter.GlobalSections.UserShippingDetailsSection.UnMerge()
                        '--------------------------------------------------------------------------
                        ' Fetching new header from data base
                        '--------------------------------------------------------------------------
                        Dim _UserShippingDetails As CT.Data.Interfaces.UserShippingDetailsInterface = Nothing
                        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                            _UserShippingDetails = New CT.Data.VehiclePlan.SevenTabs.UserShippingDetails
                        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                            _UserShippingDetails = New CT.Data.RigPlan.SevenTabs.UserShippingDetails
                        End If



                        strOutput = _UserShippingDetails.GetTndPlanHeader(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.BuildPhase, Form.DataCenter.ProgramConfig.BuildType)
                        If strOutput Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        strOutputData = _UserShippingDetails.GetPlanData(Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Form.DataCenter.ProgramConfig.BuildType)
                        If strOutputData Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

                    Case Form.DataCenter.GlobalSections.SectionName.UpdatePackSection
                        counter = Form.DataCenter.GlobalSections.UpdatePackSectionFirstColumn + 1
                        TotalColumn = counter + Form.DataCenter.GlobalSections.UpdatePackSection.Columns.Count - 3
                        CurrentHiddenStateOfSection = Form.DataCenter.GlobalSections.UpdatePackSection.Columns.Hidden
                        'Form.DataCenter.GlobalSections.UpdatePackSection.UnMerge()
                        '--------------------------------------------------------------------------
                        ' Fetching new header from data base
                        '--------------------------------------------------------------------------
                        Dim _Updatepack As CT.Data.Interfaces.UpdatepackInterface = Nothing
                        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                            _Updatepack = New CT.Data.VehiclePlan.SevenTabs.Updatepack
                        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                            _Updatepack = New CT.Data.RigPlan.SevenTabs.Updatepack
                        End If

                        strOutput = _Updatepack.GetTndPlanHeader(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.BuildPhase, Form.DataCenter.ProgramConfig.BuildType)
                        If strOutput Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        strOutputData = _Updatepack.GetPlanData(Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Form.DataCenter.ProgramConfig.BuildType)
                        If strOutputData Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

                End Select


                If TotalColumn = 0 And counter = 0 Then Exit Sub


                'For i As Integer = counter To TotalColumn
                '    Form.DataCenter.GlobalValues.WS.Columns(i).delete
                '    Form.DataCenter.GlobalValues.WS.Columns(TotalColumn + 1).Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Type.Missing)
                'Next


                FisrtOfRange = Form.DataCenter.GlobalValues.WS.Cells(1, counter)
                EndOfRange = Form.DataCenter.GlobalValues.WS.Cells(4, TotalColumn)
                all = Form.DataCenter.GlobalValues.WS.Range(FisrtOfRange, EndOfRange)
                all.ClearContents()
                all.Interior.Color = Excel.Constants.xlNone
                all.UnMerge()

                '-----------------------------------------------
                ' add column 
                '-----------------------------------------------
                Dim i As Integer
                For i = (TotalColumn - counter + 1) To strOutput.GetUpperBound(1) - 1
                    Form.DataCenter.GlobalValues.WS.Columns(TotalColumn).Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Type.Missing)
                    Form.DataCenter.GlobalValues.WS.Columns(TotalColumn).Hidden = CurrentHiddenStateOfSection
                    'Form.DataCenter.GlobalValues.WS.Columns(TotalColumn - 1).Hidden = CurrentHiddenStateOfSection
                Next

                '-----------------------------------------------
                ' delete column 
                '-----------------------------------------------
                i = strOutput.GetUpperBound(1)
                While i < (TotalColumn - counter + 1)
                    Form.DataCenter.GlobalValues.WS.Columns(TotalColumn).delete
                    i = i + 1
                End While



                DataCenter.GlobalValues.WS.Range("A1").EntireRow.Hidden = False
                TotalColumn = counter + strOutput.GetUpperBound(1) - 1
                FisrtOfRange = Form.DataCenter.GlobalValues.WS.Cells(1, counter)
                EndOfRange = Form.DataCenter.GlobalValues.WS.Cells(4, TotalColumn)
                all = Form.DataCenter.GlobalValues.WS.Range(FisrtOfRange, EndOfRange)
                all.Value2 = strOutput
                FisrtOfRange = Form.DataCenter.GlobalValues.WS.Cells(5, counter)
                EndOfRange = Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.ProgramConfig.LastRow, TotalColumn)
                all = Form.DataCenter.GlobalValues.WS.Range(FisrtOfRange, EndOfRange)
                all.Value2 = strOutputData


                FisrtOfRange = Nothing
                EndOfRange = Nothing
                all = Nothing

                While (counter <= TotalColumn + 1)

                    obj = Form.DataCenter.GlobalValues.WS.Cells(2, counter).Value2
                    If (obj IsNot Nothing) Then

                        If (FisrtOfRange Is Nothing) Then
                            FisrtOfRange = Form.DataCenter.GlobalValues.WS.Cells(2, counter)
                        ElseIf (FisrtOfRange.Value2 = (Form.DataCenter.GlobalValues.WS.Cells(2, counter)).Value2) Then
                            EndOfRange = Form.DataCenter.GlobalValues.WS.Cells(2, counter)
                        ElseIf (FisrtOfRange IsNot Nothing And EndOfRange IsNot Nothing) Then
                            CType(Form.DataCenter.GlobalValues.WS.Range(FisrtOfRange, EndOfRange), Excel.Range).Merge(False)
                            FisrtOfRange = Form.DataCenter.GlobalValues.WS.Cells(2, counter)
                            EndOfRange = Nothing
                        Else
                            FisrtOfRange = Form.DataCenter.GlobalValues.WS.Cells(2, counter)
                            EndOfRange = Nothing
                        End If
                    End If

                    TopOfRange = Nothing
                    ButtomOfRange = Form.DataCenter.GlobalValues.WS.Cells(4, counter)

                    If ((Form.DataCenter.GlobalValues.WS.Cells(4, counter)).Value2 = (Form.DataCenter.GlobalValues.WS.Cells(3, counter)).Value2) Then
                        TopOfRange = Form.DataCenter.GlobalValues.WS.Cells(3, counter)
                        If ((Form.DataCenter.GlobalValues.WS.Cells(3, counter)).Value2 = (Form.DataCenter.GlobalValues.WS.Cells(2, counter)).Value2) Then
                            TopOfRange = Form.DataCenter.GlobalValues.WS.Cells(2, counter)
                        End If
                    End If

                    If (TopOfRange IsNot Nothing And ButtomOfRange IsNot Nothing) Then
                        CType(Form.DataCenter.GlobalValues.WS.Range(ButtomOfRange, TopOfRange), Excel.Range).Merge(False)
                        TopOfRange = Nothing
                        ButtomOfRange = Nothing
                    End If

                    Dim strColor As String = ""

                    Try
                        strColor = Form.DataCenter.GlobalValues.WS.Cells(1, counter).Value2.ToString()
                    Catch ex As Exception
                    End Try

                    Dim rgbs As String() = strColor.Split(";")
                    If (rgbs.Length = 2) Then
                        If Form.DataCenter.GlobalValues.WS.Cells(4, counter).MergeCells = True Then
                            CType(Form.DataCenter.GlobalValues.WS.Cells(4, counter), Excel.Range).MergeArea.Cells(1, 1).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(0).Substring(0, 3)), Integer.Parse(rgbs(0).Substring(3, 3)), Integer.Parse(rgbs(0).Substring(6, 3))))
                            CType(Form.DataCenter.GlobalValues.WS.Cells(4, counter), Excel.Range).MergeArea.Cells(1, 1).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(1).Substring(0, 3)), Integer.Parse(rgbs(1).Substring(3, 3)), Integer.Parse(rgbs(1).Substring(6, 3))))
                        End If

                    End If
                    counter = counter + 1
                End While

                DataCenter.GlobalValues.WS.Range("A1").EntireRow.Hidden = True
                Form.DataCenter.GlobalSections.ColorSection.EntireRow.Hidden = True
                DataCenter.GlobalValues.WS.Range("A1").EntireColumn.Hidden = True
                DataCenter.GlobalValues.WS.Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column).EntireColumn.Hidden = True
                Form.DataCenter.GlobalSections.SectionSection.RowHeight = 75
                Form.DataCenter.GlobalSections.HeaderSection.RowHeight = 23
                Form.DataCenter.GlobalSections.DescriptionSection.RowHeight = 85

            Catch ex As Exception
                'MsgBox(ex.Message)
                'System.Windows.Forms.MessageBox.Show(ex.Message, "ApplyColorAndMergeToHeaderSection (Specification Section)", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.DrawTndPlanHeader, ex.Message), "ApplyColorAndMergeToHeaderSection (Specification Section)", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
            Finally
                objGlobal.sbProtectPlan()
            End Try
        End Sub
    End Class
End Namespace