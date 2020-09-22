Imports System.Globalization
Imports Excel = Microsoft.Office.Interop.Excel
Imports Office = Microsoft.Office.Core
Imports Microsoft.Office.Tools.Excel

Namespace Form.DataCenter

    Friend NotInheritable Class GlobalValues


        Public Shared ci As New CultureInfo("en-US")
        Public Shared cal As Calendar = ci.Calendar
        Public Shared myCWR As CalendarWeekRule = vbFirstFourDays 'ci.DateTimeFormat.CalendarWeekRule

        '--------------------------------------------------------------
        ' These values must be here
        '--------------------------------------------------------------
        'Public Const vbEmptyColor As Long = 16777215
        Public Const Mname As String = "MyPopUpMenu"
        Public Const ConstPwd As String = "admin123"
        '--------------------------------------------------------------
        Public Shared wsEve As Form.DisplayUtilities.clsWorksheetEvents

        Public Shared objWBCurrent As Microsoft.Office.Interop.Excel.Workbook
        Public Shared objCurrWorksheet As Microsoft.Office.Tools.Excel.Worksheet
        'Public Shared NativeWorksheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing

        Public Shared Property strUserPermissionLevel() As String
            Get
                Try
                    If Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B47").Value = "" Then
                        Return Nothing
                    Else
                        Return Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B47").Value
                    End If
                Catch ex As Exception
                    Return Nothing
                End Try
            End Get
            Set(ByVal value As String)
                Try
                    Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B47").Value = value
                Catch
                End Try
            End Set
        End Property
        Public Shared Property intProgValue() As Integer
            Get
                Return Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B46").Value)
            End Get
            Set(ByVal value As Integer)
                Try
                    Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B46").Value = value
                Catch ex As Exception

                End Try

            End Set
        End Property
        Public Shared Property bolfrmwidth() As Boolean
            Get
                Return If(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B45").Value = "True", True, False)
            End Get
            Set(ByVal value As Boolean)
                Try
                    Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B45").Value = If(value.ToString = "True", True, False)
                Catch ex As Exception

                End Try

            End Set
        End Property


        Public Shared Property TotalRow() As Integer
            Get
                Try
                    Return Integer.Parse(Form.DataCenter.GlobalValues.WS.Parent.worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B30").Value)
                Catch ex As Exception
                    Return 0
                End Try
            End Get
            Set(ByVal value As Integer)
                Try
                    Form.DataCenter.GlobalValues.WS.Parent.worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B30").Value = value
                Catch ex As Exception

                End Try

            End Set
        End Property
        Public Shared Property CurrentTotalMessages() As Integer
            Get
                Try
                    Return Integer.Parse(Form.DataCenter.GlobalValues.WS.Parent.worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B50").Value)
                Catch ex As Exception
                    Return 0
                End Try
            End Get
            Set(ByVal value As Integer)
                Try
                    Form.DataCenter.GlobalValues.WS.Parent.worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B50").Value = value
                Catch ex As Exception
                End Try
            End Set
        End Property

        Public Shared Property bolPlanIsLoading() As Boolean
            Get
                Try
                    Return Boolean.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B31").Value)
                Catch ex As Exception
                    Return False
                End Try
            End Get
            Set(ByVal value As Boolean)
                Try
                    Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B31").Value = value
                Catch ex As Exception
                End Try
            End Set
        End Property
        Public Shared Property bolPlanDrawInProgress() As Boolean
            Get
                Try
                    Return Boolean.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B32").Value)
                Catch ex As Exception
                    Return False
                End Try
            End Get
            Set(ByVal value As Boolean)
                Try
                    Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B32").Value = value
                Catch ex As Exception

                End Try

            End Set
        End Property
        Public Shared Property bolCutCopyMode() As Boolean
            Get
                Try
                    Return Boolean.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B33").Value)
                Catch ex As Exception
                    Return False
                End Try
            End Get
            Set(ByVal value As Boolean)
                Try
                    Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B33").Value = value
                Catch ex As Exception

                End Try

            End Set
        End Property
        Public Shared Property bolInsertMode() As Boolean
            Get
                Try
                    Return Boolean.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B34").Value)
                Catch ex As Exception
                    Return False
                End Try
            End Get
            Set(ByVal value As Boolean)
                Try
                    Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B34").Value = value
                Catch ex As Exception

                End Try

            End Set
        End Property
        Public Shared Property strCutAddress() As String
            Get
                Try
                    Return Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B35").Value
                Catch ex As Exception
                    Return False
                End Try
            End Get
            Set(ByVal value As String)
                Try
                    Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B35").Value = value
                Catch ex As Exception

                End Try

            End Set
        End Property
        Public Shared Property bolUserCaseSelected() As Boolean
            Get
                Try
                    Return Boolean.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B36").Value)
                Catch ex As Exception
                    Return False
                End Try
            End Get
            Set(ByVal value As Boolean)
                Try
                    Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B36").Value = value
                Catch ex As Exception

                End Try

            End Set
        End Property
        Public Shared Property strUserCaseSelected() As String
            Get
                Return Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B37").Value
            End Get
            Set(ByVal value As String)
                Try
                    Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B37").Value = value
                Catch ex As Exception

                End Try
            End Set
        End Property
        Public Shared Property strEditAddress() As String
            Get
                Return Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B38").Value
            End Get
            Set(ByVal value As String)
                Try
                    Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B38").Value = value
                Catch ex As Exception

                End Try

            End Set
        End Property
        Public Shared Property strSelAllAddress() As String
            Get
                Return Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B39").Value
            End Get
            Set(ByVal value As String)
                Try
                    Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B39").Value = value
                Catch ex As Exception

                End Try

            End Set
        End Property
        Public Shared Property bolSelAll() As Boolean
            Get
                Try
                    Return Boolean.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B40").Value)
                Catch ex As Exception
                    Return False
                End Try

            End Get
            Set(ByVal value As Boolean)
                Try
                    Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B40").Value = value
                Catch ex As Exception

                End Try

            End Set
        End Property
        Public Shared Property strCopyAddress() As String
            Get
                Return Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B41").Value
            End Get
            Set(ByVal value As String)
                Try
                    Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B41").Value = value
                Catch ex As Exception

                End Try

            End Set
        End Property
        Public Shared Property bolRefreshCompleted() As Boolean
            Get
                Try
                    Return Boolean.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B42").Value)
                Catch ex As Exception
                    Return False
                End Try
            End Get
            Set(ByVal value As Boolean)
                Try
                    Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B42").Value = value
                Catch ex As Exception

                End Try

            End Set
        End Property
        Public Shared Property bolCut() As Boolean
            Get
                Try
                    Return Boolean.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B43").Value)
                Catch ex As Exception
                    Return False
                End Try
            End Get
            Set(ByVal value As Boolean)
                Try
                    Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B43").Value = value
                Catch ex As Exception
                End Try
            End Set
        End Property
        Public Shared Property bolCopy() As Boolean
            Get
                Try
                    Return Boolean.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B44").Value)
                Catch ex As Exception
                    Return False
                End Try
            End Get
            Set(ByVal value As Boolean)
                Try
                    Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString).Range("B44").Value = value
                Catch ex As Exception

                End Try

            End Set
        End Property
        Public Shared ReadOnly Property WS As Microsoft.Office.Tools.Excel.Worksheet
            Get
                Dim bolWasSU As Boolean
                Dim bolWasEE As Boolean

                bolWasSU = Globals.ThisAddIn.Application.ScreenUpdating
                bolWasEE = Globals.ThisAddIn.Application.EnableEvents

                Try

                    Globals.ThisAddIn.Application.EnableEvents = False
                    Globals.ThisAddIn.Application.ScreenUpdating = False

                    Dim WB As Excel.Workbook

                    If Globals.ThisAddIn.Application.ActiveWorkbook.Name Like "TndTemplate*" Then
                        Globals.ThisAddIn.Application.Worksheets(Form.DataCenter.WorkSheet.TnDPlan.ToString()).Activate
                    ElseIf objWBCurrent IsNot Nothing Then
                        For Each WB In Globals.ThisAddIn.Application.Workbooks
                            If WB.Name Like "TndTemplate*" Then
                                WB.Activate()
                                WB.Worksheets(Form.DataCenter.WorkSheet.TnDPlan.ToString).activate()
                            End If
                        Next
                    End If

                    If objWBCurrent IsNot Nothing Then
                        objWBCurrent.Activate()
                        objWBCurrent.Worksheets(Form.DataCenter.WorkSheet.TnDPlan.ToString).activate()
                    End If

                    If objCurrWorksheet Is Nothing Then
                        objCurrWorksheet = Globals.Factory.GetVstoObject(CType(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet, Excel.Worksheet))
                        objWBCurrent = CType(Globals.ThisAddIn.Application.ActiveWorkbook, Excel.Workbook)
                        objWBCurrent.Activate()
                        objWBCurrent.Worksheets(Form.DataCenter.WorkSheet.TnDPlan.ToString).activate()
                        objCurrWorksheet = Globals.Factory.GetVstoObject(CType(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet, Excel.Worksheet))
                    End If

                    If bolPlanIsLoading Then
                        WS = objCurrWorksheet
                    Else
                        WS = Globals.Factory.GetVstoObject(CType(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet, Excel.Worksheet))
                        objWBCurrent = CType(Globals.ThisAddIn.Application.ActiveWorkbook, Excel.Workbook)
                        objCurrWorksheet = Globals.Factory.GetVstoObject(CType(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet, Excel.Worksheet))
                    End If

                Catch ex As Exception
                    WS = Nothing
                    If objCurrWorksheet IsNot Nothing Then
                        If objCurrWorksheet.Name.ToString.ToLower = Form.DataCenter.WorkSheet.TnDPlan.ToString.ToLower Then
                            WS = objCurrWorksheet
                        End If
                    End If
                End Try

                Globals.ThisAddIn.Application.EnableEvents = bolWasEE
                Globals.ThisAddIn.Application.ScreenUpdating = bolWasSU

            End Get
        End Property

        Public Shared ReadOnly Property ChangeLogWs As Microsoft.Office.Tools.Excel.Worksheet
            Get
                Try
                    Globals.ThisAddIn.Application.Worksheets(Form.DataCenter.WorkSheet.ChangeLogs.ToString()).Activate
                    ChangeLogWs = Globals.Factory.GetVstoObject(CType(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet, Excel.Worksheet))
                    'wsEve = New Form.DisplayUtilities.clsWorksheetEvents
                Catch ex As Exception
                    ChangeLogWs = Nothing
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property VehicleReportWs As Microsoft.Office.Tools.Excel.Worksheet
            Get
                Try
                    Globals.ThisAddIn.Application.Worksheets(Form.DataCenter.WorkSheet.VehicleReportTemplate.ToString()).Visible = Excel.XlSheetVisibility.xlSheetVisible
                    Globals.ThisAddIn.Application.Worksheets(Form.DataCenter.WorkSheet.VehicleReportTemplate.ToString()).Activate
                    VehicleReportWs = Globals.Factory.GetVstoObject(CType(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet, Excel.Worksheet))
                    'wsEve = New Form.DisplayUtilities.clsWorksheetEvents
                Catch ex As Exception
                    VehicleReportWs = Nothing
                End Try
            End Get
        End Property

        Public WSruntime As Object = Nothing
        Public Shared TPEditProcessStepTaskPane As Microsoft.Office.Tools.CustomTaskPane = Nothing
        Public colWeekends As New Collection
        Public Shared sEditId As String

        Public Shared Sub Clear()
            TotalRow = 0
            TPEditProcessStepTaskPane = Nothing
            bolRefreshCompleted = False
        End Sub
    End Class

End Namespace
