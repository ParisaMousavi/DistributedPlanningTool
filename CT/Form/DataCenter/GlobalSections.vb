
Namespace Form.DataCenter

    Public NotInheritable Class GlobalSections
        ''' <summary>
        ''' This enum refer to the name of the sections or namedrange in Excel interface
        ''' </summary>
        Public Enum SectionName As Integer

            ColorSection = 100
            SectionSection = 101
            HeaderSection = 102
            DescriptionSection = 103

            VehicleProgramInfoSection = 1
            InstrumentationSection = 2
            NonMfcSpecificationSection = 3
            MfcSpecificationSection = 4
            ProgramInformationSection = 5
            FurtherBasicInformationSection = 6
            UserShippingDetailsSection = 7
            UpdatePackSection = 8
            TimelineSection = 9

        End Enum


        Private Shared _ErrorMessage As String
        Public Shared ReadOnly Property ErrorMessage() As String
            Get
                Return _ErrorMessage
            End Get
        End Property


        Public Shared Function IsDateValid(DateToValidate As Date) As Boolean
            Dim FindColumn As Excel.Range = Nothing

            Try
                FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(DateToValidate.ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                If FindColumn IsNot Nothing Then
                    IsDateValid = True
                Else
                    IsDateValid = False
                End If
            Catch ex As Exception
                IsDateValid = False
            End Try

        End Function




        Public Shared Function GetCellOfDate(DateToValidate As Date) As Excel.Range
            Dim FindColumn As Excel.Range = Nothing

            Try
                FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(DateToValidate.ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                GetCellOfDate = FindColumn
            Catch ex As Exception
                GetCellOfDate = Nothing
            End Try

        End Function


        ''' <summary>
        ''' This List is used in ApplyColorAndMergeToHeaderSection. 
        ''' </summary>
        Public Shared SectionFlags As New List(Of String) From {"Program Start", "Program End", "Instrumentation Start", "Instrumentation End", "Non MFC Specification Start", "Non MFC Specification End", "MFC Specification Start", "MFC Specification End", "Program Information Start", "Program Information End", "User Shipping Details Start", "User Shipping Details End", "Update Pack Start", "Update Pack End", "Further Basic Information Start", "Further Basic Information End", "Timeline Start", "Timeline End"}


        Public Shared ReadOnly Property ColorSection As Excel.Range

            Get
                Try
                    ColorSection = CType(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).Range(Form.DataCenter.GlobalSections.SectionName.ColorSection.ToString), Excel.Range)
                Catch ex As Exception
                    ColorSection = Nothing
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property ColorSectionFirstColumn As Integer
            Get
                Try
                    ColorSectionFirstColumn = CType(Form.DataCenter.GlobalSections.ColorSection.Cells(1, 1), Excel.Range).Column
                Catch ex As Exception
                    ColorSectionFirstColumn = 0
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property ColorSectionLastColumn As Integer
            Get
                Try

                    Dim rng As Excel.Range = Form.DataCenter.GlobalSections.ColorSection
                    ColorSectionLastColumn = CType(rng.Cells(1, rng.Columns.Count), Excel.Range).Column

                Catch ex As Exception
                    ColorSectionLastColumn = 0
                End Try

            End Get
        End Property



        Public Shared ReadOnly Property SectionSection As Excel.Range

            Get
                Try
                    SectionSection = CType(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).Range(Form.DataCenter.GlobalSections.SectionName.SectionSection.ToString), Excel.Range)
                Catch ex As Exception
                    SectionSection = Nothing
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property SectionSectionFirstColumn As Integer
            Get
                Try
                    SectionSectionFirstColumn = CType(Form.DataCenter.GlobalSections.SectionSection.Cells(1, 1), Excel.Range).Column
                Catch ex As Exception
                    SectionSectionFirstColumn = 0
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property SectionSectionLastColumn As Integer
            Get
                Try

                    Dim rng As Excel.Range = Form.DataCenter.GlobalSections.SectionSection
                    SectionSectionLastColumn = CType(rng.Cells(1, rng.Columns.Count), Excel.Range).Column

                Catch ex As Exception
                    SectionSectionLastColumn = 0
                End Try

            End Get
        End Property



        Public Shared ReadOnly Property HeaderSection As Excel.Range

            Get
                Try
                    HeaderSection = CType(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).Range(Form.DataCenter.GlobalSections.SectionName.HeaderSection.ToString), Excel.Range)
                Catch ex As Exception
                    HeaderSection = Nothing
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property HeaderSectionFirstColumn As Integer
            Get
                Try
                    HeaderSectionFirstColumn = CType(Form.DataCenter.GlobalSections.HeaderSection.Cells(1, 1), Excel.Range).Column
                Catch ex As Exception
                    HeaderSectionFirstColumn = 0
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property HeaderSectionLastColumn As Integer
            Get
                Try

                    Dim rng As Excel.Range = Form.DataCenter.GlobalSections.HeaderSection
                    HeaderSectionLastColumn = CType(rng.Cells(1, rng.Columns.Count), Excel.Range).Column

                Catch ex As Exception
                    HeaderSectionLastColumn = 0
                End Try

            End Get
        End Property



        Public Shared ReadOnly Property DescriptionSection As Excel.Range

            Get
                Try
                    DescriptionSection = CType(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).Range(Form.DataCenter.GlobalSections.SectionName.DescriptionSection.ToString), Excel.Range)
                Catch ex As Exception
                    DescriptionSection = Nothing
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property DescriptionSectionFirstColumn As Integer
            Get
                Try
                    DescriptionSectionFirstColumn = CType(Form.DataCenter.GlobalSections.DescriptionSection.Cells(1, 1), Excel.Range).Column
                Catch ex As Exception
                    DescriptionSectionFirstColumn = 0
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property DescriptionSectionLastColumn As Integer
            Get
                Try

                    Dim rng As Excel.Range = Form.DataCenter.GlobalSections.DescriptionSection
                    DescriptionSectionLastColumn = CType(rng.Cells(1, rng.Columns.Count), Excel.Range).Column

                Catch ex As Exception
                    DescriptionSectionLastColumn = 0
                End Try

            End Get
        End Property



        Public Shared ReadOnly Property VehicleProgramInfoSection As Excel.Range

            Get
                Try
                    VehicleProgramInfoSection = CType(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).Range(Form.DataCenter.GlobalSections.SectionName.VehicleProgramInfoSection.ToString), Excel.Range)
                Catch ex As Exception
                    VehicleProgramInfoSection = Nothing
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property VehicleProgramInfoSectionFirstColumn As Integer
            Get
                Try
                    VehicleProgramInfoSectionFirstColumn = CType(Form.DataCenter.GlobalSections.VehicleProgramInfoSection.Cells(1, 1), Excel.Range).Column
                Catch ex As Exception
                    VehicleProgramInfoSectionFirstColumn = 0
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property VehicleProgramInfoSectionLastColumn As Integer
            Get
                Try

                    Dim rng As Excel.Range = Form.DataCenter.GlobalSections.VehicleProgramInfoSection
                    VehicleProgramInfoSectionLastColumn = CType(rng.Cells(1, rng.Columns.Count), Excel.Range).Column

                Catch ex As Exception
                    VehicleProgramInfoSectionLastColumn = 0
                End Try

            End Get
        End Property




        Public Shared ReadOnly Property InstrumentationSection As Excel.Range

            Get
                Try
                    InstrumentationSection = CType(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).Range(Form.DataCenter.GlobalSections.SectionName.InstrumentationSection.ToString), Excel.Range)
                Catch ex As Exception
                    InstrumentationSection = Nothing
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property InstrumentationSectionFirstColumn As Integer
            Get
                Try
                    InstrumentationSectionFirstColumn = CType(Form.DataCenter.GlobalSections.InstrumentationSection.Cells(1, 1), Excel.Range).Column
                Catch ex As Exception
                    InstrumentationSectionFirstColumn = 0
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property InstrumentationSectionLastColumn As Integer
            Get
                Try

                    Dim rng As Excel.Range = Form.DataCenter.GlobalSections.InstrumentationSection
                    InstrumentationSectionLastColumn = CType(rng.Cells(1, rng.Columns.Count), Excel.Range).Column

                Catch ex As Exception
                    InstrumentationSectionLastColumn = 0
                End Try

            End Get
        End Property





        Public Shared ReadOnly Property NonMfcSpecificationSection As Excel.Range

            Get
                Try
                    NonMfcSpecificationSection = CType(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).Range(Form.DataCenter.GlobalSections.SectionName.NonMfcSpecificationSection.ToString), Excel.Range)
                Catch ex As Exception
                    NonMfcSpecificationSection = Nothing
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property NonMfSpecificationSectionFirstColumn As Integer
            Get
                Try
                    NonMfSpecificationSectionFirstColumn = CType(Form.DataCenter.GlobalSections.NonMfcSpecificationSection.Cells(1, 1), Excel.Range).Column
                Catch ex As Exception
                    NonMfSpecificationSectionFirstColumn = 0
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property NonMfSpecificationSectionLastColumn As Integer
            Get
                Try

                    Dim rng As Excel.Range = Form.DataCenter.GlobalSections.NonMfcSpecificationSection
                    NonMfSpecificationSectionLastColumn = CType(rng.Cells(1, rng.Columns.Count), Excel.Range).Column

                Catch ex As Exception
                    NonMfSpecificationSectionLastColumn = 0
                End Try

            End Get
        End Property

        Public Shared ReadOnly Property MfcSpecificationSection As Excel.Range

            Get
                Try
                    MfcSpecificationSection = CType(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).Range(Form.DataCenter.GlobalSections.SectionName.MfcSpecificationSection.ToString), Excel.Range)
                Catch ex As Exception
                    MfcSpecificationSection = Nothing
                End Try

            End Get
        End Property

        Public Shared ReadOnly Property MfcSpecificationSectionFirstColumn As Integer
            Get
                Try
                    MfcSpecificationSectionFirstColumn = CType(Form.DataCenter.GlobalSections.MfcSpecificationSection.Cells(1, 1), Excel.Range).Column
                Catch ex As Exception
                    MfcSpecificationSectionFirstColumn = 0
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property MfcSpecificationSectionLastColumn As Integer
            Get
                Try

                    Dim rng As Excel.Range = Form.DataCenter.GlobalSections.MfcSpecificationSection
                    MfcSpecificationSectionLastColumn = CType(rng.Cells(1, rng.Columns.Count), Excel.Range).Column

                Catch ex As Exception
                    MfcSpecificationSectionLastColumn = 0
                End Try

            End Get
        End Property




        Public Shared ReadOnly Property ProgramInformationSection As Excel.Range

            Get
                Try
                    ProgramInformationSection = CType(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).Range(Form.DataCenter.GlobalSections.SectionName.ProgramInformationSection.ToString), Excel.Range)
                Catch ex As Exception
                    ProgramInformationSection = Nothing
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property ProgramInformationSectionFirstColumn As Integer
            Get
                Try
                    ProgramInformationSectionFirstColumn = CType(Form.DataCenter.GlobalSections.ProgramInformationSection.Cells(1, 1), Excel.Range).Column
                Catch ex As Exception
                    ProgramInformationSectionFirstColumn = 0
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property ProgramInformationSectionLastColumn As Integer
            Get
                Try

                    Dim rng As Excel.Range = Form.DataCenter.GlobalSections.ProgramInformationSection
                    ProgramInformationSectionLastColumn = CType(rng.Cells(1, rng.Columns.Count), Excel.Range).Column

                Catch ex As Exception
                    ProgramInformationSectionLastColumn = 0
                End Try

            End Get
        End Property

        Public Shared ReadOnly Property FurtherBasicInformationSection As Excel.Range

            Get
                Try
                    FurtherBasicInformationSection = CType(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).Range(Form.DataCenter.GlobalSections.SectionName.FurtherBasicInformationSection.ToString), Excel.Range)
                Catch ex As Exception
                    FurtherBasicInformationSection = Nothing
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property FurtherBasicInformationSectionFirstColumn As Integer
            Get
                Try
                    FurtherBasicInformationSectionFirstColumn = CType(Form.DataCenter.GlobalSections.FurtherBasicInformationSection.Cells(1, 1), Excel.Range).Column
                Catch ex As Exception
                    FurtherBasicInformationSectionFirstColumn = 0
                End Try

            End Get
        End Property


        Public Shared ReadOnly Property FurtherBasicInformationSectionLastColumn As Integer
            Get
                Try
                    Dim rng As Excel.Range = Form.DataCenter.GlobalSections.FurtherBasicInformationSection
                    FurtherBasicInformationSectionLastColumn = CType(rng.Cells(1, rng.Columns.Count), Excel.Range).Column

                Catch ex As Exception
                    FurtherBasicInformationSectionLastColumn = 0
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property UserShippingDetailsSection As Excel.Range
            Get
                Try
                    UserShippingDetailsSection = CType(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).Range(Form.DataCenter.GlobalSections.SectionName.UserShippingDetailsSection.ToString), Excel.Range)
                Catch ex As Exception
                    UserShippingDetailsSection = Nothing
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property UserShippingDetailsSectionFirstColumn As Integer
            Get
                Try
                    UserShippingDetailsSectionFirstColumn = CType(Form.DataCenter.GlobalSections.UserShippingDetailsSection.Cells(1, 1), Excel.Range).Column
                Catch ex As Exception
                    UserShippingDetailsSectionFirstColumn = 0
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property UserShippingDetailsSectionLastColumn As Integer
            Get
                Try
                    Dim rng As Excel.Range = Form.DataCenter.GlobalSections.UserShippingDetailsSection
                    UserShippingDetailsSectionLastColumn = CType(rng.Cells(1, rng.Columns.Count), Excel.Range).Column
                Catch ex As Exception
                    UserShippingDetailsSectionLastColumn = 0
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property UpdatePackSection As Excel.Range
            Get
                Try
                    UpdatePackSection = CType(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).Range(Form.DataCenter.GlobalSections.SectionName.UpdatePackSection.ToString), Excel.Range)
                Catch ex As Exception
                    UpdatePackSection = Nothing
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property UpdatePackSectionFirstColumn As Integer
            Get
                Try
                    UpdatePackSectionFirstColumn = CType(Form.DataCenter.GlobalSections.UpdatePackSection.Cells(1, 1), Excel.Range).Column
                Catch ex As Exception
                    UpdatePackSectionFirstColumn = 0
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property UpdatePackSectionLastColumn As Integer
            Get
                Try
                    Dim rng As Excel.Range = Form.DataCenter.GlobalSections.UpdatePackSection
                    UpdatePackSectionLastColumn = CType(rng.Cells(1, rng.Columns.Count), Excel.Range).Column
                Catch ex As Exception
                    UpdatePackSectionLastColumn = 0
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property TimeLineSection As Excel.Range
            Get
                Try
                    TimeLineSection = CType(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).Range(Form.DataCenter.GlobalSections.SectionName.TimelineSection.ToString), Excel.Range)
                Catch ex As Exception
                    TimeLineSection = Nothing
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property TimeLineSectionFirstColumn As Integer
            Get
                Try
                    TimeLineSectionFirstColumn = CType(Form.DataCenter.GlobalSections.TimeLineSection.Cells(1, 1), Excel.Range).Column
                Catch ex As Exception
                    TimeLineSectionFirstColumn = 0
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property TimeLineSectionLastColumn As Integer
            Get
                Try
                    Dim rng As Excel.Range = Form.DataCenter.GlobalSections.TimeLineSection
                    TimeLineSectionLastColumn = CType(rng.Cells(1, rng.Columns.Count), Excel.Range).Column
                Catch ex As Exception
                    TimeLineSectionLastColumn = 0
                End Try
            End Get
        End Property

    End Class
End Namespace
