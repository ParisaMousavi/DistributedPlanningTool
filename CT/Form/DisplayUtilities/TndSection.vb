
Imports System.Windows.Forms


Namespace Form.DisplayUtilities
    Friend NotInheritable Class TndSection
        ''' <summary>
        ''' Detecting the first 4 Rows for formating
        ''' Color , Section, Header, Description
        ''' </summary>
        Public Shared Sub DetectFirstElementarySections(TotalColumn As Integer)


            Dim cell As Excel.Range = Nothing
            Dim borders As Excel.Borders = Nothing


            Try


                '-------------------- Color Section ------------------------
                cell = Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.WS.Cells(1, 1), Form.DataCenter.GlobalValues.WS.Cells(1, TotalColumn))

                Form.DataCenter.GlobalValues.WS.Controls.AddNamedRange(cell, Form.DataCenter.GlobalSections.SectionName.ColorSection.ToString())


                '-------------------- Section Section ------------------------
                cell = Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.WS.Cells(2, 1), Form.DataCenter.GlobalValues.WS.Cells(2, TotalColumn))

                Form.DataCenter.GlobalValues.WS.Controls.AddNamedRange(cell, Form.DataCenter.GlobalSections.SectionName.SectionSection.ToString())

                '-------------------- Hedear Section ------------------------
                cell = Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.WS.Cells(3, 1), Form.DataCenter.GlobalValues.WS.Cells(3, TotalColumn))

                Form.DataCenter.GlobalValues.WS.Controls.AddNamedRange(cell, Form.DataCenter.GlobalSections.SectionName.HeaderSection.ToString())

                '-------------------- Description Section ------------------------
                cell = Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.WS.Cells(4, 1), Form.DataCenter.GlobalValues.WS.Cells(4, TotalColumn))

                Form.DataCenter.GlobalValues.WS.Controls.AddNamedRange(cell, Form.DataCenter.GlobalSections.SectionName.DescriptionSection.ToString())

                '--------------------------------------------------------------------------------------------------------
                '  "Program Section"
                '--------------------------------------------------------------------------------------------------------

                cell = Nothing

                cell = DisplayUtilities.Utilities.FindRange("Program Start", "Program End")
                If (cell IsNot Nothing) Then




                    DataCenter.GlobalValues.WS.Controls.AddNamedRange(cell, DataCenter.GlobalSections.SectionName.VehicleProgramInfoSection.ToString())

                    cell.Rows(3).Style = Style.Styles.TnsStyleName.VehicleProgramInfoStyle.ToString()


                    DataCenter.GlobalSections.VehicleProgramInfoSection.ColumnWidth = 7D
                    borders = DataCenter.GlobalSections.VehicleProgramInfoSection.Borders
                    borders.LineStyle = Excel.XlLineStyle.xlContinuous
                    borders.Weight = 1D

                    'set seprators
                    cell.Columns(1).Style = Style.Styles.TnsStyleName.StartEndColumnsStyle.ToString()
                    cell.Columns(cell.Columns.Count).Style = Style.Styles.TnsStyleName.StartEndColumnsStyle.ToString()

                    cell.Columns(1).ColumnWidth = 2D
                    cell.Columns(cell.Columns.Count).ColumnWidth = 2D




                End If


                '--------------------------------------------------------------------------------------------------------
                '  "TimeLine  Section"
                '--------------------------------------------------------------------------------------------------------
                cell = Nothing
                cell = DisplayUtilities.Utilities.FindRange("Timeline Start", "Timeline End")
                If (cell IsNot Nothing) Then


                    'Style in Template
                    'Style.Styles.AddTimelineStyle() 

                    'DataCenter.GlobalSections.TimeLineSection = 
                    DataCenter.GlobalValues.WS.Controls.AddNamedRange(
                        cell, DataCenter.GlobalSections.SectionName.TimelineSection.ToString())


                    'aplying font style to first row because of merge
                    cell.Rows(3).Style = Style.Styles.TnsStyleName.TimelineStyle.ToString()
                    cell.Rows(1).Style = Style.Styles.TnsStyleName.TimelineStyle.ToString()



                    'StartupTnD.DataCenter.clsSections.TimeLineSection.Style = StartupTnD.Style.Styles.TnsStyleName.TimelineStyle.ToString() 
                    DataCenter.GlobalSections.TimeLineSection.ColumnWidth = 2
                    borders = DataCenter.GlobalSections.TimeLineSection.Borders
                    borders.LineStyle = Excel.XlLineStyle.xlContinuous
                    borders.Weight = 1D

                    'set seprators
                    cell.Columns(1).Style = Style.Styles.TnsStyleName.StartEndColumnsStyle.ToString()
                    cell.Columns(cell.Columns.Count).Style = Style.Styles.TnsStyleName.StartEndColumnsStyle.ToString()

                    cell.Columns(1).ColumnWidth = 2D
                    cell.Columns(cell.Columns.Count).ColumnWidth = 2D

                End If


                '---------------------------------------------------------------
                ' This exception is here to not to create and look for the other sections in Generic plan
                '---------------------------------------------------------------
                If Form.DataCenter.ProgramConfig.IsGeneric = True Then Throw New Exception("000")

                '--------------------------------------------------------------------------------------------------------
                '  "Instrumentation  Section"
                '--------------------------------------------------------------------------------------------------------
                cell = Nothing
                cell = DisplayUtilities.Utilities.FindRange("Instrumentation Start", "Instrumentation End")
                If (cell IsNot Nothing) Then



                    DataCenter.GlobalValues.WS.Controls.AddNamedRange(cell, DataCenter.GlobalSections.SectionName.InstrumentationSection.ToString())


                    cell.Rows(1).Style = Style.Styles.TnsStyleName.InstrumentationSectionStyle.ToString()
                    cell.Rows(2).Style = Style.Styles.TnsStyleName.InstrumentationHeaderStyle.ToString()

                    DataCenter.GlobalSections.InstrumentationSection.ColumnWidth = 7D
                    borders = DataCenter.GlobalSections.InstrumentationSection.Borders
                    borders.LineStyle = Excel.XlLineStyle.xlContinuous
                    borders.Weight = 1D

                    'set seprators
                    cell.Columns(1).Style = Style.Styles.TnsStyleName.StartEndColumnsStyle.ToString()
                    cell.Columns(cell.Columns.Count).Style = Style.Styles.TnsStyleName.StartEndColumnsStyle.ToString()

                    cell.Columns(1).ColumnWidth = 2D
                    cell.Columns(cell.Columns.Count).ColumnWidth = 2D


                End If


                '--------------------------------------------------------------------------------------------------------
                '  "Non Mfc Specification Section"
                '--------------------------------------------------------------------------------------------------------
                cell = Nothing
                cell = DisplayUtilities.Utilities.FindRange("Non MFC Specification Start", "Non MFC Specification End")
                If (cell IsNot Nothing) Then


                    'Style in Template
                    'Style.Styles.AddNonMfcSpecificationHeaderStyle() 

                    DataCenter.GlobalValues.WS.Controls.AddNamedRange(cell, DataCenter.GlobalSections.SectionName.NonMfcSpecificationSection.ToString())

                    cell.Rows(1).Style = Style.Styles.TnsStyleName.NonMfcSpecificationHeaderStyle.ToString()

                    DataCenter.GlobalSections.NonMfcSpecificationSection.ColumnWidth = 5D
                    borders = DataCenter.GlobalSections.NonMfcSpecificationSection.Borders
                    borders.LineStyle = Excel.XlLineStyle.xlContinuous
                    borders.Weight = 1D

                    'set seprators
                    cell.Columns(1).Style = Style.Styles.TnsStyleName.StartEndColumnsStyle.ToString()
                    cell.Columns(cell.Columns.Count).Style = Style.Styles.TnsStyleName.StartEndColumnsStyle.ToString()

                    cell.Columns(1).ColumnWidth = 2D
                    cell.Columns(cell.Columns.Count).ColumnWidth = 2D


                End If



                '--------------------------------------------------------------------------------------------------------
                '  "Mfc Specification Section"
                '--------------------------------------------------------------------------------------------------------

                cell = Nothing
                cell = DisplayUtilities.Utilities.FindRange("MFC Specification Start", "MFC Specification End")
                If (cell IsNot Nothing) Then



                    DataCenter.GlobalValues.WS.Controls.AddNamedRange(cell, DataCenter.GlobalSections.SectionName.MfcSpecificationSection.ToString())


                    cell.Rows(1).Style = Style.Styles.TnsStyleName.MfcSpecificationSectionStyle.ToString()
                    cell.Rows(2).Style = Style.Styles.TnsStyleName.MfcSpecificationHeaderStyle.ToString()
                    cell.Rows(3).Style = Style.Styles.TnsStyleName.MfcSpecificationDescriptionStyle.ToString()


                    DataCenter.GlobalSections.MfcSpecificationSection.ColumnWidth = 7D
                    borders = DataCenter.GlobalSections.MfcSpecificationSection.Borders
                    borders.LineStyle = Excel.XlLineStyle.xlContinuous
                    borders.Weight = 1D

                    'set seprators
                    cell.Columns(1).Style = Style.Styles.TnsStyleName.StartEndColumnsStyle.ToString()
                    cell.Columns(cell.Columns.Count).Style = Style.Styles.TnsStyleName.StartEndColumnsStyle.ToString()

                    cell.Columns(1).ColumnWidth = 2D
                    cell.Columns(cell.Columns.Count).ColumnWidth = 2D


                End If



                '--------------------------------------------------------------------------------------------------------
                '  "Program Information Section"
                '--------------------------------------------------------------------------------------------------------

                cell = Nothing
                'finding section here 
                cell = DisplayUtilities.Utilities.FindRange("Program Information Start", "Program Information End")
                If (cell IsNot Nothing) Then


                    ' adding style if the section Is found.
                    'Style in Template
                    'Style.Styles.AddProgramInformationStyle() 

                    'making named range for the section
                    DataCenter.GlobalValues.WS.Controls.AddNamedRange(cell, DataCenter.GlobalSections.SectionName.ProgramInformationSection.ToString())

                    'aplying font style to first row because of merge
                    cell.Rows(1).Style = Style.Styles.TnsStyleName.ProgramInformationStyle.ToString()

                    'applying the size of interface object
                    DataCenter.GlobalSections.ProgramInformationSection.ColumnWidth = 10D
                    borders = DataCenter.GlobalSections.ProgramInformationSection.Borders
                    borders.LineStyle = Excel.XlLineStyle.xlContinuous
                    borders.Weight = 1D

                    'set seprators
                    cell.Columns(1).Style = Style.Styles.TnsStyleName.StartEndColumnsStyle.ToString()
                    cell.Columns(cell.Columns.Count).Style = Style.Styles.TnsStyleName.StartEndColumnsStyle.ToString()

                    cell.Columns(1).ColumnWidth = 2D
                    cell.Columns(cell.Columns.Count).ColumnWidth = 2D



                End If



                '--------------------------------------------------------------------------------------------------------
                '  "Further Basic Information Section"
                '--------------------------------------------------------------------------------------------------------

                cell = Nothing
                cell = DisplayUtilities.Utilities.FindRange("Further Basic Information Start", "Further Basic Information End")
                If (cell IsNot Nothing) Then

                    ' adding style if the section Is found.
                    'Style in Template
                    'Style.Styles.AddFurtherBasicInformationStyle() 

                    'making named range for the section
                    DataCenter.GlobalValues.WS.Controls.AddNamedRange(cell, DataCenter.GlobalSections.SectionName.FurtherBasicInformationSection.ToString())

                    'aplying font style to first row because of merge
                    cell.Rows(1).Style = Style.Styles.TnsStyleName.FurtherBasicInformationStyle.ToString()


                    'applying the size of interface object
                    DataCenter.GlobalSections.FurtherBasicInformationSection.ColumnWidth = 10D
                    borders = DataCenter.GlobalSections.FurtherBasicInformationSection.Borders
                    borders.LineStyle = Excel.XlLineStyle.xlContinuous
                    borders.Weight = 1D


                    'set seprators
                    cell.Columns(1).Style = Style.Styles.TnsStyleName.StartEndColumnsStyle.ToString()
                    cell.Columns(cell.Columns.Count).Style = Style.Styles.TnsStyleName.StartEndColumnsStyle.ToString()

                    cell.Columns(1).ColumnWidth = 2D
                    cell.Columns(cell.Columns.Count).ColumnWidth = 2D



                End If



                '--------------------------------------------------------------------------------------------------------
                '  "User Shipping Details Section"
                '--------------------------------------------------------------------------------------------------------

                cell = Nothing
                cell = DisplayUtilities.Utilities.FindRange("User Shipping Details Start", "User Shipping Details End")
                If (cell IsNot Nothing) Then

                    ' adding style if the section Is found.
                    'Style in Template
                    'Style.Styles.AddUserShippingDetailsStyle() 

                    'making named range for the section
                    DataCenter.GlobalValues.WS.Controls.AddNamedRange(cell, DataCenter.GlobalSections.SectionName.UserShippingDetailsSection.ToString())

                    'aplying font style to first row because of merge
                    cell.Rows(1).Style = Style.Styles.TnsStyleName.UserShippingDetailsStyle.ToString()

                    'applying the size of interface object
                    DataCenter.GlobalSections.UserShippingDetailsSection.ColumnWidth = 10D
                    borders = DataCenter.GlobalSections.UserShippingDetailsSection.Borders
                    borders.LineStyle = Excel.XlLineStyle.xlContinuous
                    borders.Weight = 1D

                    'set seprators
                    cell.Columns(1).Style = Style.Styles.TnsStyleName.StartEndColumnsStyle.ToString()
                    cell.Columns(cell.Columns.Count).Style = Style.Styles.TnsStyleName.StartEndColumnsStyle.ToString()

                    cell.Columns(1).ColumnWidth = 2D
                    cell.Columns(cell.Columns.Count).ColumnWidth = 2D



                End If



                '--------------------------------------------------------------------------------------------------------
                '  "Update Pack Section"
                '--------------------------------------------------------------------------------------------------------

                cell = Nothing
                cell = DisplayUtilities.Utilities.FindRange("Update Pack Start", "Update Pack End")
                If (cell IsNot Nothing) Then

                    DataCenter.GlobalValues.WS.Controls.AddNamedRange(cell, DataCenter.GlobalSections.SectionName.UpdatePackSection.ToString())


                    cell.Rows(1).Style = Style.Styles.TnsStyleName.TimelineStyle.ToString()
                    cell.Rows(2).Style = Style.Styles.TnsStyleName.TimelineStyle.ToString()


                    DataCenter.GlobalSections.UpdatePackSection.ColumnWidth = 10 '2
                    borders = DataCenter.GlobalSections.UpdatePackSection.Borders
                    borders.LineStyle = Excel.XlLineStyle.xlContinuous
                    borders.Weight = 1D



                    'set seprators
                    cell.Columns(1).Style = Style.Styles.TnsStyleName.StartEndColumnsStyle.ToString()
                    cell.Columns(cell.Columns.Count).Style = Style.Styles.TnsStyleName.StartEndColumnsStyle.ToString()

                    cell.Columns(1).ColumnWidth = 2D
                    cell.Columns(cell.Columns.Count).ColumnWidth = 2D


                End If


            Catch ex As Exception

                'If ex.Message <> "000" Then System.Windows.Forms.MessageBox.Show(ex.Message, "DetectFirstElementarySections", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                If ex.Message <> "000" Then MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.TndSection, ex.Message), "DetectFirstElementarySections", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
            End Try
        End Sub


        Public Shared Sub CleanRange(all As Excel.Range)
            all.Clear()
            all.ClearContents()
            all.ClearFormats()
            all.UnMerge()
            all.Style = Style.Styles.TnsStyleName.ProcessStepStyle.ToString()
        End Sub
        Public Shared Sub CleanNumbersAfterRefreshSection(rng As Excel.Range)
            Dim row As Excel.Range

            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic


            Try
                With Form.DataCenter.GlobalValues.WS
                    For Each row In rng.Rows
                        Globals.ThisAddIn.Application.ScreenUpdating = False
                        Globals.ThisAddIn.Application.EnableEvents = False
                        Globals.ThisAddIn.Application.DisplayAlerts = False
                        Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic

                        .Cells(row.Row, Form.DataCenter.GlobalSections.InstrumentationSectionFirstColumn).value2 = ""
                        .Cells(row.Row, Form.DataCenter.GlobalSections.InstrumentationSectionLastColumn).value2 = ""

                        .Cells(row.Row, Form.DataCenter.GlobalSections.InstrumentationSectionFirstColumn).interior.color = .Cells(4, Form.DataCenter.GlobalSections.InstrumentationSectionFirstColumn).interior.color
                        .Cells(row.Row, Form.DataCenter.GlobalSections.InstrumentationSectionLastColumn).interior.color = .Cells(4, Form.DataCenter.GlobalSections.InstrumentationSectionLastColumn).interior.color

                        .Cells(row.Row, Form.DataCenter.GlobalSections.MfcSpecificationSectionFirstColumn).value2 = ""
                        .Cells(row.Row, Form.DataCenter.GlobalSections.MfcSpecificationSectionLastColumn).value2 = ""

                        .Cells(row.Row, Form.DataCenter.GlobalSections.MfcSpecificationSectionFirstColumn).interior.color = .Cells(4, Form.DataCenter.GlobalSections.MfcSpecificationSectionFirstColumn).interior.color
                        .Cells(row.Row, Form.DataCenter.GlobalSections.MfcSpecificationSectionLastColumn).interior.color = .Cells(4, Form.DataCenter.GlobalSections.MfcSpecificationSectionLastColumn).interior.color


                        .Cells(row.Row, Form.DataCenter.GlobalSections.NonMfSpecificationSectionFirstColumn).value2 = ""
                        .Cells(row.Row, Form.DataCenter.GlobalSections.NonMfSpecificationSectionLastColumn).value2 = ""

                        .Cells(row.Row, Form.DataCenter.GlobalSections.NonMfSpecificationSectionFirstColumn).interior.color = .Cells(4, Form.DataCenter.GlobalSections.NonMfSpecificationSectionFirstColumn).interior.color
                        .Cells(row.Row, Form.DataCenter.GlobalSections.NonMfSpecificationSectionLastColumn).interior.color = .Cells(4, Form.DataCenter.GlobalSections.NonMfSpecificationSectionLastColumn).interior.color


                        .Cells(row.Row, Form.DataCenter.GlobalSections.UpdatePackSectionFirstColumn).value2 = ""
                        .Cells(row.Row, Form.DataCenter.GlobalSections.UpdatePackSectionLastColumn).value2 = ""

                        .Cells(row.Row, Form.DataCenter.GlobalSections.UpdatePackSectionFirstColumn).interior.color = .Cells(4, Form.DataCenter.GlobalSections.UpdatePackSectionFirstColumn).interior.color
                        .Cells(row.Row, Form.DataCenter.GlobalSections.UpdatePackSectionLastColumn).interior.color = .Cells(4, Form.DataCenter.GlobalSections.UpdatePackSectionLastColumn).interior.color

                        .Cells(row.Row, Form.DataCenter.GlobalSections.UserShippingDetailsSectionFirstColumn).value2 = ""
                        .Cells(row.Row, Form.DataCenter.GlobalSections.UserShippingDetailsSectionLastColumn).value2 = ""

                        .Cells(row.Row, Form.DataCenter.GlobalSections.UserShippingDetailsSectionFirstColumn).interior.color = .Cells(4, Form.DataCenter.GlobalSections.UserShippingDetailsSectionFirstColumn).interior.color
                        .Cells(row.Row, Form.DataCenter.GlobalSections.UserShippingDetailsSectionLastColumn).interior.color = .Cells(4, Form.DataCenter.GlobalSections.UserShippingDetailsSectionLastColumn).interior.color

                        .Cells(row.Row, Form.DataCenter.GlobalSections.VehicleProgramInfoSectionFirstColumn).value2 = ""
                        .Cells(row.Row, Form.DataCenter.GlobalSections.VehicleProgramInfoSectionLastColumn).value2 = ""

                        .Cells(row.Row, Form.DataCenter.GlobalSections.VehicleProgramInfoSectionFirstColumn).interior.color = .Cells(4, Form.DataCenter.GlobalSections.VehicleProgramInfoSectionFirstColumn).interior.color
                        .Cells(row.Row, Form.DataCenter.GlobalSections.VehicleProgramInfoSectionLastColumn).interior.color = .Cells(4, Form.DataCenter.GlobalSections.VehicleProgramInfoSectionLastColumn).interior.color

                        .Cells(row.Row, Form.DataCenter.GlobalSections.ProgramInformationSectionFirstColumn).value2 = ""
                        .Cells(row.Row, Form.DataCenter.GlobalSections.ProgramInformationSectionLastColumn).value2 = ""

                        .Cells(row.Row, Form.DataCenter.GlobalSections.ProgramInformationSectionFirstColumn).interior.color = .Cells(4, Form.DataCenter.GlobalSections.ProgramInformationSectionFirstColumn).interior.color
                        .Cells(row.Row, Form.DataCenter.GlobalSections.ProgramInformationSectionLastColumn).interior.color = .Cells(4, Form.DataCenter.GlobalSections.ProgramInformationSectionLastColumn).interior.color

                        .Cells(row.Row, Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn).value2 = ""
                        .Cells(row.Row, Form.DataCenter.GlobalSections.TimeLineSectionLastColumn).value2 = ""

                        .Cells(row.Row, Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn).interior.color = .Cells(4, Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn).interior.color
                        .Cells(row.Row, Form.DataCenter.GlobalSections.TimeLineSectionLastColumn).interior.color = .Cells(4, Form.DataCenter.GlobalSections.TimeLineSectionLastColumn).interior.color

                        .Cells(row.Row, Form.DataCenter.GlobalSections.FurtherBasicInformationSectionFirstColumn).value2 = ""
                        .Cells(row.Row, Form.DataCenter.GlobalSections.FurtherBasicInformationSectionLastColumn).value2 = ""

                        .Cells(row.Row, Form.DataCenter.GlobalSections.FurtherBasicInformationSectionFirstColumn).interior.color = .Cells(4, Form.DataCenter.GlobalSections.FurtherBasicInformationSectionFirstColumn).interior.color
                        .Cells(row.Row, Form.DataCenter.GlobalSections.FurtherBasicInformationSectionLastColumn).interior.color = .Cells(4, Form.DataCenter.GlobalSections.FurtherBasicInformationSectionLastColumn).interior.color

                    Next
                End With
            Catch ex As Exception
            End Try
        End Sub
    End Class
End Namespace