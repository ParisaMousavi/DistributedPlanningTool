Imports System.Data
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms

Namespace Form.Reports
    Public Class EngineTransmissionReport
        Dim colEngineData As New List(Of String)()
        Dim colTransData = New List(Of String)()
        Public Class clsEngTrans
            Public strEngine As String
            Public strTrans As String
            Public intVehicles As Integer
            Public intBucks As Integer
            Public intRigs As Integer
            Public intRebuilds As Integer
            Public StrType As String
        End Class
        Public Sub EngineTransmissionReport(HCID As Integer, Region As String)
            Try
                Dim WS1 = Globals.Factory.GetVstoObject(CType(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet, Excel.Worksheet))
                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                GetETData(Region)
                Dim rst As New System.Data.DataTable
                Dim Wb As Workbook
                Dim ws As Worksheet
                Dim intCnt As Integer
                Dim strDisplay As String
                Dim bolEngineStarted As Boolean
                Dim bolTransStarted As Boolean
                Dim colEngine As New Dictionary(Of String, Object)
                Dim colTrans As New Dictionary(Of String, Object)
                Dim sTnDRegion As String(,) = Nothing
                Dim intCol As Integer
                Dim rngBorder As Range
                Dim rngMid As Range
                Dim BuildType As String = Form.DataCenter.ProgramConfig.BuildType

                'Dim _plan As CT.Data.VehiclePlan.Plan = New CT.Data.VehiclePlan.Plan()
                Dim _PlanInterface As Data.Interfaces.PlanInterface

                If BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                    _PlanInterface = New Data.VehiclePlan.Plan
                ElseIf BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                    _PlanInterface = New Data.RigPlan.Plan
                Else
                    Exit Sub
                End If

                Dim Engine As New Dictionary(Of String, String)
                Dim trans As New Dictionary(Of String, String)

                rst = _PlanInterface.GetCTEnginesAndTransmissions(HCID, BuildType)
                For Each row As DataRow In rst.Rows
                    For Each colmn As DataColumn In rst.Columns
                        If bolEngineStarted Then
                            If Not colmn.ColumnName = "EngineEnd - TransStart" Then
                                Engine.Add(colmn.Ordinal, colmn.ColumnName)
                            End If
                        End If
                        If bolTransStarted Then
                            If Not colmn.ColumnName = "TransEnd" Then
                                trans.Add(colmn.Ordinal, colmn.ColumnName)
                            End If
                        End If
                        If colmn.ColumnName = "EngineStart" Then bolEngineStarted = True
                        If colmn.ColumnName = "EngineEnd - TransStart" Then
                            bolTransStarted = True
                            bolEngineStarted = False
                        End If
                    Next
                    Exit For
                Next

                Wb = Globals.ThisAddIn.Application.Workbooks.Add
                ws = Wb.Worksheets(1)
                strDisplay = "Engine and Transmission counts report" & vbLf
                strDisplay = strDisplay & "HCID - " & rst.Rows(0)("HealthChartId") 'rst.rst!HealthChartId & vbLf
                strDisplay = strDisplay & "Program description - " & rst.Rows(0)("ProgramDescription") ' rst!ProgramDescription & vbLf
                strDisplay = strDisplay & "Build Phase - " & rst.Rows(0)("BuildPhase") ' rst!BuildPhase

                For i = 0 To Engine.Count - 1
                    Dim objEngine = New clsEngTrans()
                    For j = 0 To rst.Rows.Count - 1
                        Dim key As Integer = Engine.ElementAt(i).Key
                        objEngine.strEngine = Engine.ElementAt(i).Value
                        If rst.Rows(j)(key).ToString() = "" Then
                            rst.Rows(j)(key) = 0
                        End If
                        If rst.Rows(j)("BuildTypes") = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                            objEngine.intVehicles = rst.Rows(j)(key)
                        ElseIf rst.Rows(j)("BuildTypes") = CT.Data.DataCenter.BuildType.Buck.ToString() Then
                            objEngine.intBucks = rst.Rows(j)(key)
                        ElseIf rst.Rows(j)("BuildTypes") = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                            objEngine.intRigs = rst.Rows(j)(key)
                        ElseIf rst.Rows(j)("BuildTypes") = CT.Data.DataCenter.BuildType.Rebuild.ToString() Then
                            objEngine.intRebuilds = rst.Rows(j)(key)
                        End If
                    Next
                    Dim isExist As Integer = -1
                    isExist = colEngineData.IndexOf(objEngine.strEngine)
                    If isExist > 0 Then
                        objEngine.StrType = Strings.Split(colEngineData.Item(isExist), "~")(0)
                    End If
                    colEngine.Add(objEngine.strEngine, objEngine)
                Next

                For i = 0 To trans.Count - 1
                    Dim objTrans = New clsEngTrans()
                    For j = 0 To rst.Rows.Count - 1
                        Dim key As Integer = trans.ElementAt(i).Key
                        objTrans.strTrans = trans.ElementAt(i).Value
                        If rst.Rows(j)(key).ToString() = "" Then
                            rst.Rows(j)(key) = 0
                        End If
                        If rst.Rows(j)("BuildTypes") = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                            objTrans.intVehicles = rst.Rows(j)(key)
                        ElseIf rst.Rows(j)("BuildTypes") = CT.Data.DataCenter.BuildType.Buck.ToString() Then
                            objTrans.intBucks = rst.Rows(j)(key)
                        ElseIf rst.Rows(j)("BuildTypes") = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                            objTrans.intRigs = rst.Rows(j)(key)
                        ElseIf rst.Rows(j)("BuildTypes") = CT.Data.DataCenter.BuildType.Rebuild.ToString() Then
                            objTrans.intRebuilds = rst.Rows(j)(key)
                        End If

                    Next
                    Dim isExist As Integer = -1
                    isExist = colTransData.IndexOf(objTrans.strTrans)
                    If isExist > 0 Then
                        objTrans.StrType = Strings.Split(colTransData.Item(isExist), "~")(0)
                    End If
                    colTrans.Add(objTrans.strTrans, objTrans)
                Next


                Dim colEngineXCC = New Dictionary(Of String, Object)
                Dim colTransXCC = New Dictionary(Of String, Object)

                bolEngineStarted = False
                bolTransStarted = False

                Dim EngineXCC As New Dictionary(Of String, String)
                Dim transXCC As New Dictionary(Of String, String)

                ' Dim _plan As CT.Data.Plan = New CT.Data.Plan()
                rst = _PlanInterface.GetXCCEnginesAndTransmissions(HCID, BuildType)
                For Each row As DataRow In rst.Rows
                    For Each colmn As DataColumn In rst.Columns
                        If bolEngineStarted Then
                            If Not colmn.ColumnName = "EngineEnd - TransStart" Then
                                EngineXCC.Add(colmn.Ordinal, colmn.ColumnName)
                            End If
                        End If
                        If bolTransStarted Then
                            If Not colmn.ColumnName = "TransEnd" Then
                                transXCC.Add(colmn.Ordinal, colmn.ColumnName)
                            End If
                        End If
                        If colmn.ColumnName = "EngineStart" Then bolEngineStarted = True
                        If colmn.ColumnName = "EngineEnd - TransStart" Then
                            bolTransStarted = True
                            bolEngineStarted = False
                        End If
                    Next
                    Exit For
                Next


                For i = 0 To EngineXCC.Count - 1
                    Dim objEngine = New clsEngTrans()
                    For j = 0 To rst.Rows.Count - 1
                        Dim key As Integer = EngineXCC.ElementAt(i).Key
                        objEngine.strEngine = EngineXCC.ElementAt(i).Value
                        If rst.Rows(j)(key).ToString() = "" Then
                            rst.Rows(j)(key) = 0
                        End If
                        If rst.Rows(j)("pe29_BuildTypeCode") = "Vehicle" Then
                            objEngine.intVehicles = rst.Rows(j)(key)
                        ElseIf rst.Rows(j)("pe29_BuildTypeCode") = "Buck" Then
                            objEngine.intBucks = IIf(rst.Rows(j)(key) Is Nothing, Nothing, rst.Rows(j)(key))
                        ElseIf rst.Rows(j)("pe29_BuildTypeCode") = "Rig" Then
                            objEngine.intRigs = rst.Rows(j)(key)
                        ElseIf rst.Rows(j)("pe29_BuildTypeCode") = "Rebuild" Then
                            objEngine.intRebuilds = rst.Rows(j)(key)
                        End If
                    Next
                    Dim isExist As Integer = -1
                    isExist = colEngineData.IndexOf(objEngine.strEngine)
                    If isExist > 0 Then
                        objEngine.StrType = Strings.Split(colEngineData.Item(isExist), "~")(0)
                    End If
                    colEngineXCC.Add(objEngine.strEngine, objEngine)
                Next

                For i = 0 To transXCC.Count - 1
                    Dim objTrans = New clsEngTrans()
                    For j = 0 To rst.Rows.Count - 1
                        Dim key As Integer = transXCC.ElementAt(i).Key
                        objTrans.strTrans = transXCC.ElementAt(i).Value
                        If rst.Rows(j)(key).ToString() = "" Then
                            rst.Rows(j)(key) = 0
                        End If
                        If rst.Rows(j)("pe29_BuildTypeCode") = "Vehicle" Then
                            objTrans.intVehicles = rst.Rows(j)(key)
                        ElseIf rst.Rows(j)("pe29_BuildTypeCode") = "Buck" Then
                            objTrans.intBucks = rst.Rows(j)(key)
                        ElseIf rst.Rows(j)("pe29_BuildTypeCode") = "Rig" Then
                            objTrans.intRigs = rst.Rows(j)(key)
                        ElseIf rst.Rows(j)("pe29_BuildTypeCode") = "Rebuild" Then
                            objTrans.intRebuilds = rst.Rows(j)(key)
                        End If

                    Next
                    Dim isExist As Integer = -1
                    isExist = colTransData.IndexOf(objTrans.strTrans)
                    If isExist > 0 Then
                        objTrans.StrType = Strings.Split(colTransData.Item(isExist), "~")(0)
                    End If
                    colTransXCC.Add(objTrans.strTrans, objTrans)
                Next



                Dim colHeaderE As New Dictionary(Of Object, String)
                Dim colHeaderT As New Dictionary(Of Object, String)


                For intCnt = 0 To colEngine.Count - 1
                    Dim objEngine = New clsEngTrans()
                    objEngine = colEngine.ElementAt(intCnt).Value
                    If Not colHeaderE.ContainsValue(objEngine.strEngine) Then
                        colHeaderE.Add(objEngine, colHeaderE.Count + 1)
                    End If

                Next

                For intCnt = 0 To colEngineXCC.Count - 1
                    Dim objEngine = New clsEngTrans()
                    objEngine = colEngineXCC.ElementAt(intCnt).Value
                    If Not colHeaderE.ContainsValue(objEngine.strEngine) Then
                        colHeaderE.Add(objEngine, colHeaderE.Count + 1)
                    End If
                Next


                For intCnt = 0 To colTrans.Count - 1
                    Dim objTrans = New clsEngTrans()
                    objTrans = colTrans.ElementAt(intCnt).Value
                    If Not colHeaderT.ContainsValue(objTrans.strTrans) Then
                        colHeaderT.Add(objTrans, colHeaderT.Count + 1)
                    End If
                Next

                For intCnt = 0 To colTransXCC.Count - 1
                    Dim objTrans = New clsEngTrans()
                    objTrans = colTransXCC.ElementAt(intCnt).Value
                    If Not colHeaderT.ContainsValue(objTrans.strTrans) Then
                        colHeaderT.Add(objTrans, colHeaderT.Count + 1)
                    End If
                Next
                'colHeaderE = colEngineData.Distinct().ToList()
                'colHeaderE = colEngineXCC.Distinct().ToList()
                'colHeaderT = colTrans.Distinct().ToList()
                'colHeaderT = colTransXCC.Distinct().ToList()
                ' colHeaderT.

                'colHeaderT.Sort()
                'colHeaderE.Sort()
                intCol = 10
                Globals.ThisAddIn.Application.DisplayAlerts = False
                Dim rngEngine As Range
                Dim rngTrans As Range
                Dim rngGas As Range
                Dim rngDiesel As Range
                Dim rngManual As Range
                Dim rngAuto As Range

                With ws
                    With .Range("C15:C18")
                        .Value = "CT"
                        .Orientation = 90
                        .Font.Bold = True
                        .VerticalAlignment = Constants.xlCenter
                        .HorizontalAlignment = Constants.xlCenter
                        .Merge()
                    End With
                    With .Range("C20:C23")
                        .Value = "XCC"
                        .Orientation = 90
                        .Font.Bold = True
                        .VerticalAlignment = Constants.xlCenter
                        .HorizontalAlignment = Constants.xlCenter
                        .Merge()
                    End With

                    With .Range("D15:H15")
                        .Value = "Vehicles"
                        .HorizontalAlignment = Constants.xlRight
                        .Font.Bold = True
                        .Merge()
                    End With
                    With .Range("D16:H16")
                        .Value = "Bucks"
                        .Font.Bold = True
                        .HorizontalAlignment = Constants.xlRight
                        .Merge()
                    End With
                    With .Range("D17:H17")
                        .Value = "Rigs"
                        .Font.Bold = True
                        .HorizontalAlignment = Constants.xlRight
                        .Merge()
                    End With
                    With .Range("D18:H18")
                        .Value = "Rebuilds"
                        .Font.Bold = True
                        .HorizontalAlignment = Constants.xlRight
                        .Merge()
                    End With

                    With .Range("D20:H20")
                        .Value = "Vehicles"
                        .HorizontalAlignment = Constants.xlRight
                        .Font.Bold = True
                        .Merge()
                    End With
                    With .Range("D21:H21")
                        .Value = "Bucks"
                        .HorizontalAlignment = Constants.xlRight
                        .Font.Bold = True
                        .Merge()
                    End With
                    With .Range("D22:H22")
                        .Value = "Rigs"
                        .HorizontalAlignment = Constants.xlRight
                        .Font.Bold = True
                        .Merge()
                    End With
                    With .Range("D23:H23")
                        .Value = "Rebuilds"
                        .HorizontalAlignment = Constants.xlRight
                        .Font.Bold = True
                        .Merge()
                    End With

                    Dim colEnginechk As New Dictionary(Of String, Object)
                    For intCnt = 0 To colHeaderE.Count - 1
                        Dim objEngine = New clsEngTrans()
                        objEngine = colHeaderE.ElementAt(intCnt).Key
                        If Not colEnginechk.ContainsKey(objEngine.strEngine) Then
                            .Cells(2, intCol) = "Engine"
                            If Not rngEngine Is Nothing Then
                                rngEngine = Globals.ThisAddIn.Application.Union(rngEngine, .Cells(2, intCol))
                            Else
                                rngEngine = .Cells(2, intCol)
                            End If
                            If objEngine.StrType = "Gas" Then
                                If Not rngGas Is Nothing Then
                                    rngGas = Globals.ThisAddIn.Application.Union(rngGas, .Cells(3, intCol))
                                Else
                                    rngGas = .Cells(3, intCol)
                                End If
                            Else
                                If Not rngDiesel Is Nothing Then
                                    rngDiesel = Globals.ThisAddIn.Application.Union(rngDiesel, .Cells(3, intCol))
                                Else
                                    rngDiesel = .Cells(3, intCol)
                                End If
                            End If
                            .Cells(3, intCol) = objEngine.StrType
                            .Range(.Cells(2, intCol), .Cells(3, intCol)).Font.Bold = True
                            .Range(.Cells(4, intCol), .Cells(13, intCol)).Value = objEngine.strEngine
                            .Range(.Cells(4, intCol), .Cells(13, intCol)).Merge()
                            .Range(.Cells(4, intCol), .Cells(13, intCol)).Orientation = 90
                            If colEngine.ContainsKey(objEngine.strEngine) Then
                                objEngine = colEngine(objEngine.strEngine)
                                .Cells(15, intCol).value2 = IIf(objEngine.intVehicles = 0, "", objEngine.intVehicles)
                                .Cells(16, intCol).value2 = IIf(objEngine.intBucks = 0, "", objEngine.intBucks)
                                .Cells(17, intCol).value2 = IIf(objEngine.intRigs = 0, "", objEngine.intRigs)
                                .Cells(18, intCol).value2 = IIf(objEngine.intRebuilds = 0, "", objEngine.intRebuilds)
                                colEnginechk.Add(objEngine.strEngine, "")
                            End If

                            .Range(.Cells(15, intCol), .Cells(18, intCol)).HorizontalAlignment = Constants.xlCenter
                            .Range(.Cells(15, intCol), .Cells(18, intCol)).VerticalAlignment = Constants.xlCenter
                            intCol = intCol + 1
                        End If
                    Next
                    colEnginechk.Clear()
                    intCol = intCol + 2

                    Dim colTranschk As New Dictionary(Of String, Object)
                    For intCnt = 0 To colHeaderT.Count - 1
                        Dim objTrans = New clsEngTrans()
                        objTrans = colHeaderT.ElementAt(intCnt).Key
                        If Not colTranschk.ContainsKey(objTrans.strTrans) Then
                            .Cells(2, intCol) = "Transmission"
                            .Cells(3, intCol) = Strings.Replace(objTrans.StrType, "Automatic", "Auto")
                            If Not rngTrans Is Nothing Then
                                rngTrans = Globals.ThisAddIn.Application.Union(rngTrans, .Cells(2, intCol))
                            Else
                                rngTrans = .Cells(2, intCol)
                            End If
                            If objTrans.StrType = "Manual" Then
                                If Not rngManual Is Nothing Then
                                    rngManual = Globals.ThisAddIn.Application.Union(rngManual, .Cells(3, intCol))
                                Else
                                    rngManual = .Cells(3, intCol)
                                End If
                            Else
                                If Not rngAuto Is Nothing Then
                                    rngAuto = Globals.ThisAddIn.Application.Union(rngAuto, .Cells(3, intCol))
                                Else
                                    rngAuto = .Cells(3, intCol)
                                End If
                            End If
                            .Range(.Cells(2, intCol), .Cells(3, intCol)).Font.Bold = True
                            .Range(.Cells(4, intCol), .Cells(13, intCol)).Value = objTrans.strTrans
                            .Range(.Cells(4, intCol), .Cells(13, intCol)).Merge()
                            .Range(.Cells(4, intCol), .Cells(13, intCol)).Orientation = 90
                            If colTrans.ContainsKey(objTrans.strTrans) Then

                                objTrans = colTrans(objTrans.strTrans)
                                .Cells(15, intCol) = IIf(objTrans.intVehicles = 0, "", objTrans.intVehicles)
                                .Cells(16, intCol) = IIf(objTrans.intBucks = 0, "", objTrans.intBucks)
                                .Cells(17, intCol) = IIf(objTrans.intRigs = 0, "", objTrans.intRigs)
                                .Cells(18, intCol) = IIf(objTrans.intRebuilds = 0, "", objTrans.intRebuilds)
                                colTranschk.Add(objTrans.strTrans, "")
                            End If

                            .Range(.Cells(15, intCol), .Cells(18, intCol)).HorizontalAlignment = Constants.xlCenter
                            .Range(.Cells(15, intCol), .Cells(18, intCol)).VerticalAlignment = Constants.xlCenter
                            intCol = intCol + 1
                        End If
                    Next
                    colTranschk.Clear()
                    intCol = 10

                    rngBorder = Globals.ThisAddIn.Application.Union(rngEngine, .Range(.Cells(rngEngine.Row, 10), .Cells(13, rngEngine.Columns.Count + 9)))

                    rngBorder.BorderAround(XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium)
                    rngBorder.Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
                    rngBorder.Borders(XlBordersIndex.xlInsideHorizontal).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    rngBorder.Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
                    rngBorder.Borders(XlBordersIndex.xlInsideVertical).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    rngBorder.Font.Size = 8
                    rngBorder.HorizontalAlignment = Constants.xlCenter

                    With rngBorder.Interior
                        .Pattern = Constants.xlSolid
                        .PatternColorIndex = Constants.xlAutomatic
                        .ThemeColor = XlThemeColor.xlThemeColorDark1
                        .TintAndShade = -0.249977111117893
                        .PatternTintAndShade = 0
                    End With

                    rngBorder = Globals.ThisAddIn.Application.Union(rngTrans, .Range(.Cells(rngTrans.Row, rngEngine.Columns.Count + 12), .Cells(13, rngTrans.Columns.Count + rngEngine.Columns.Count + 11)))

                    rngBorder.BorderAround(XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium)
                    rngBorder.Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
                    rngBorder.Borders(XlBordersIndex.xlInsideHorizontal).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    rngBorder.Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
                    rngBorder.Borders(XlBordersIndex.xlInsideVertical).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    rngBorder.Font.Size = 8
                    rngBorder.HorizontalAlignment = Constants.xlCenter


                    With rngBorder.Interior
                        .Pattern = Constants.xlSolid
                        .PatternColorIndex = Constants.xlAutomatic
                        .ThemeColor = XlThemeColor.xlThemeColorDark1
                        .TintAndShade = -0.249977111117893
                        .PatternTintAndShade = 0
                    End With

                    If Not rngEngine Is Nothing Then
                        rngEngine.Font.Size = 8
                        rngEngine.VerticalAlignment = Constants.xlCenter
                        rngEngine.Merge()
                        rngEngine.HorizontalAlignment = Constants.xlCenter
                    End If

                    If Not rngTrans Is Nothing Then
                        rngTrans.Font.Size = 8
                        rngTrans.VerticalAlignment = Constants.xlCenter
                        rngTrans.Merge()
                        rngTrans.HorizontalAlignment = Constants.xlCenter
                    End If

                    If Not rngGas Is Nothing Then
                        rngGas.Font.Size = 8
                        rngGas.VerticalAlignment = Constants.xlCenter
                        rngGas.Merge()
                        rngGas.HorizontalAlignment = Constants.xlCenter
                    End If
                    If Not rngDiesel Is Nothing Then
                        rngDiesel.Font.Size = 8
                        rngDiesel.VerticalAlignment = Constants.xlCenter
                        rngDiesel.Merge()
                        rngDiesel.HorizontalAlignment = Constants.xlCenter

                    End If
                    If Not rngManual Is Nothing Then
                        rngManual.Font.Size = 8
                        rngManual.VerticalAlignment = Constants.xlCenter
                        rngManual.Merge()
                        rngManual.HorizontalAlignment = Constants.xlCenter
                    End If
                    If Not rngAuto Is Nothing Then
                        rngAuto.Font.Size = 8
                        rngAuto.VerticalAlignment = Constants.xlCenter
                        rngAuto.Merge()
                        rngAuto.HorizontalAlignment = Constants.xlCenter
                    End If

                    rngEngine = Nothing
                    rngTrans = Nothing
                    Dim rngECT As Range
                    Dim rngTCT As Range

                    For intCnt = 0 To colHeaderE.Count - 1
                        Dim objEngine = New clsEngTrans()
                        objEngine = colHeaderE.ElementAt(intCnt).Key
                        If Not colEnginechk.ContainsKey(objEngine.strEngine) Then
                            If Not rngEngine Is Nothing Then
                                rngEngine = Globals.ThisAddIn.Application.Union(rngEngine, .Range(.Cells(20, intCol), .Cells(23, intCol)))
                            Else
                                rngEngine = .Range(.Cells(20, intCol), .Cells(23, intCol))
                            End If
                            If Not rngECT Is Nothing Then
                                rngECT = Globals.ThisAddIn.Application.Union(rngECT, .Range(.Cells(15, intCol), .Cells(18, intCol)))
                            Else
                                rngECT = .Range(.Cells(15, intCol), .Cells(18, intCol))
                            End If
                            Dim value = .Cells(4, intCol).Formula
                            Dim k As Integer = -1
                            If colEngineXCC.ContainsKey(value) Then
                                For Each item In colEngineXCC.Keys
                                    If item = value Then
                                        k += 1
                                        Exit For
                                    End If
                                    k += 1
                                Next
                                '     If colContains(colEngineXCC, .Cells(4, intCol)) Then
                                objEngine = colEngineXCC.ElementAt(k).Value

                                .Cells(20, intCol) = IIf(objEngine.intVehicles = 0, "", objEngine.intVehicles)
                                .Cells(21, intCol) = IIf(objEngine.intBucks = 0, "", objEngine.intBucks)
                                .Cells(22, intCol) = IIf(objEngine.intRigs = 0, "", objEngine.intRigs)
                                .Cells(23, intCol) = IIf(objEngine.intRebuilds = 0, "", objEngine.intRebuilds)
                                colEnginechk.Add(objEngine.strEngine, "")

                            End If
                            If Not .Cells(20, intCol) Is .Cells(15, intCol) Then
                                .Cells(15, intCol).Interior.Color = 12040422 : .Cells(20, intCol).Interior.Color = 12040422
                            Else
                                .Cells(15, intCol).Interior.Color = 12379352 : .Cells(20, intCol).Interior.Color = 12379352
                            End If
                            If Not .Cells(21, intCol) Is .Cells(16, intCol) Then
                                .Cells(16, intCol).Interior.Color = 12040422 : .Cells(21, intCol).Interior.Color = 12040422
                            Else
                                .Cells(16, intCol).Interior.Color = 12379352 : .Cells(21, intCol).Interior.Color = 12379352
                            End If
                            If Not .Cells(22, intCol) Is .Cells(17, intCol) Then
                                .Cells(17, intCol).Interior.Color = 12040422 : .Cells(22, intCol).Interior.Color = 12040422
                            Else
                                .Cells(17, intCol).Interior.Color = 12379352 : .Cells(22, intCol).Interior.Color = 12379352
                            End If
                            If Not .Cells(23, intCol) Is .Cells(18, intCol) Then
                                .Cells(18, intCol).Interior.Color = 12040422 : .Cells(23, intCol).Interior.Color = 12040422
                            Else
                                .Cells(18, intCol).Interior.Color = 12379352 : .Cells(23, intCol).Interior.Color = 12379352
                            End If
                            .Range(.Cells(20, intCol), .Cells(23, intCol)).HorizontalAlignment = Constants.xlCenter
                            .Range(.Cells(20, intCol), .Cells(23, intCol)).VerticalAlignment = Constants.xlCenter
                            intCol = intCol + 1
                        End If
                    Next

                    rngEngine.BorderAround(XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium)
                    rngEngine.Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
                    rngEngine.Borders(XlBordersIndex.xlInsideHorizontal).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    rngEngine.Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
                    rngEngine.Borders(XlBordersIndex.xlInsideVertical).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin

                    rngECT.BorderAround(XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium)
                    rngECT.Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
                    rngECT.Borders(XlBordersIndex.xlInsideHorizontal).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    rngECT.Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
                    rngECT.Borders(XlBordersIndex.xlInsideVertical).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin

                    intCol = intCol + 2

                    rngMid = .Cells(20, intCol - 1)

                    For intCnt = 0 To colHeaderT.Count - 1
                        Dim objTrans = New clsEngTrans()
                        objTrans = colHeaderT.ElementAt(intCnt).Key
                        If Not colTranschk.ContainsKey(objTrans.strTrans) Then
                            If Not rngTrans Is Nothing Then
                                rngTrans = Globals.ThisAddIn.Application.Union(rngTrans, .Range(.Cells(20, intCol), .Cells(23, intCol)))
                            Else
                                rngTrans = .Range(.Cells(20, intCol), .Cells(23, intCol))
                            End If
                            If Not rngTCT Is Nothing Then
                                rngTCT = Globals.ThisAddIn.Application.Union(rngTCT, .Range(.Cells(15, intCol), .Cells(18, intCol)))
                            Else
                                rngTCT = .Range(.Cells(15, intCol), .Cells(18, intCol))
                            End If
                            Dim value = .Cells(4, intCol).Formula
                            Dim k As Integer = -1
                            If colTransXCC.ContainsKey(value) Then
                                For Each item In colTransXCC.Keys
                                    If item = value Then
                                        k += 1
                                        Exit For
                                    End If
                                    k += 1
                                Next

                                objTrans = colTransXCC.ElementAt(k).Value

                                .Cells(20, intCol) = IIf(objTrans.intVehicles = 0, "", objTrans.intVehicles)
                                .Cells(21, intCol) = IIf(objTrans.intBucks = 0, "", objTrans.intBucks)
                                .Cells(22, intCol) = IIf(objTrans.intRigs = 0, "", objTrans.intRigs)
                                .Cells(23, intCol) = IIf(objTrans.intRebuilds = 0, "", objTrans.intRebuilds)
                                colTranschk.Add(objTrans.strTrans, "")

                            End If
                            'If colTransXCC.ContainsKey(.Cells(4, intCol)) Then
                            ' If colContains(colTransXCC, .Cells(4, intCol)) Then


                            '  End If
                            If Not .Cells(20, intCol) Is .Cells(15, intCol) Then
                                .Cells(15, intCol).Interior.Color = 12040422 : .Cells(20, intCol).Interior.Color = 12040422
                            Else
                                .Cells(15, intCol).Interior.Color = 12379352 : .Cells(20, intCol).Interior.Color = 12379352
                            End If
                            If Not .Cells(21, intCol) Is .Cells(16, intCol) Then
                                .Cells(16, intCol).Interior.Color = 12040422 : .Cells(21, intCol).Interior.Color = 12040422
                            Else
                                .Cells(16, intCol).Interior.Color = 12379352 : .Cells(21, intCol).Interior.Color = 12379352
                            End If
                            If Not .Cells(22, intCol) Is .Cells(17, intCol) Then
                                .Cells(17, intCol).Interior.Color = 12040422 : .Cells(22, intCol).Interior.Color = 12040422
                            Else
                                .Cells(17, intCol).Interior.Color = 12379352 : .Cells(22, intCol).Interior.Color = 12379352
                            End If
                            If Not .Cells(23, intCol) Is .Cells(18, intCol) Then
                                .Cells(18, intCol).Interior.Color = 12040422 : .Cells(23, intCol).Interior.Color = 12040422
                            Else
                                .Cells(18, intCol).Interior.Color = 12379352 : .Cells(23, intCol).Interior.Color = 12379352
                            End If
                            .Range(.Cells(20, intCol), .Cells(23, intCol)).HorizontalAlignment = Constants.xlCenter
                            .Range(.Cells(20, intCol), .Cells(23, intCol)).VerticalAlignment = Constants.xlCenter
                            intCol = intCol + 1
                        End If
                    Next

                    rngTrans.BorderAround(XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium)
                    rngTrans.Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
                    rngTrans.Borders(XlBordersIndex.xlInsideHorizontal).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    rngTrans.Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
                    rngTrans.Borders(XlBordersIndex.xlInsideVertical).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin

                    rngTCT.BorderAround(XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium)
                    rngTCT.Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
                    rngTCT.Borders(XlBordersIndex.xlInsideHorizontal).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    rngTCT.Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
                    rngTCT.Borders(XlBordersIndex.xlInsideVertical).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin

                    rngECT = .Range("C15:H18")
                    rngTCT = .Range("C20:H23")

                    rngECT.BorderAround(XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium)
                    rngECT.Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
                    rngECT.Borders(XlBordersIndex.xlInsideHorizontal).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    rngECT.Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
                    rngECT.Borders(XlBordersIndex.xlInsideVertical).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin

                    rngTCT.BorderAround(XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium)
                    rngTCT.Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
                    rngTCT.Borders(XlBordersIndex.xlInsideHorizontal).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    rngTCT.Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
                    rngTCT.Borders(XlBordersIndex.xlInsideVertical).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin

                    rngECT.Font.Size = 8
                    rngTCT.Font.Size = 8

                    rngECT.VerticalAlignment = Constants.xlCenter
                    rngTCT.VerticalAlignment = Constants.xlCenter

                    rngECT.EntireColumn.ColumnWidth = 3
                    rngTCT.EntireColumn.ColumnWidth = 3

                    With rngECT.Interior
                        .Pattern = Constants.xlSolid
                        .PatternColorIndex = Constants.xlAutomatic
                        .ThemeColor = XlThemeColor.xlThemeColorDark1
                        .TintAndShade = -0.249977111117893
                        .PatternTintAndShade = 0
                    End With

                    With rngTCT.Interior
                        .Pattern = Constants.xlSolid
                        .PatternColorIndex = Constants.xlAutomatic
                        .ThemeColor = XlThemeColor.xlThemeColorDark1
                        .TintAndShade = -0.249977111117893
                        .PatternTintAndShade = 0
                    End With

                    .Cells.EntireColumn.ColumnWidth = 4
                    .Cells.EntireRow.RowHeight = 18
                    .Cells(1, "I").EntireColumn.ColumnWidth = 1.6
                    .Cells(1, "B").EntireColumn.ColumnWidth = 1.6
                    .Range("14:24").RowHeight = 12
                    .Range("15:23").RowHeight = 25
                    .Range("19:19").RowHeight = 12
                    .Cells(1, "A").EntireColumn.Delete
                End With

                Wb.Activate()

                For intCnt = 1 To 4
                    ws.Range("A1").EntireRow.Insert(XlInsertShiftDirection.xlShiftDown)
                Next

                ws.Cells(1, rngMid.Column - 1).EntireColumn.Delete
                ws.Range("B:G").EntireColumn.ColumnWidth = 3
                ws.Range("1:1").EntireRow.RowHeight = 8
                Globals.ThisAddIn.Application.ActiveWindow.DisplayZeros = False
                Globals.ThisAddIn.Application.ActiveWindow.DisplayGridlines = False

                ws.Name = "Engine & trans counts"
                rngMid.EntireColumn.ColumnWidth = 1.6

                Dim strTempTitle() As String, shp As Microsoft.Office.Interop.Excel.Shape

                strTempTitle = Strings.Split(WS1.Shapes(0).TextFrame2.TextRange.Characters.Text, vbLf)
                shp = ws.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, ws.Range("I1").Left, 5, ws.Range(ws.Range("I1"), ws.Cells(1, ws.UsedRange.Columns.Count + 1)).Width, 65)
                shp.TextFrame2.TextRange.Characters.Text = "Engine and Transmission CT DB Vs XCC DB Quantities report" & vbLf & strTempTitle(0) & vbLf & strTempTitle(1) & vbLf & strTempTitle(2)
                shp.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter
                shp.ShapeStyle = MsoShapeStyleIndex.msoShapeStylePreset16
                shp.Line.Visible = MsoTriState.msoFalse

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.EnableEvents = True
                Globals.ThisAddIn.Application_WorkbookActivate(Wb)
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message & "-" & ex.Message, "EngineTransmission Report", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
            End Try
        End Sub

        Private Sub GetETData(Region As String)

            Dim sTnDRegion As String = String.Empty
            Dim rst As New System.Data.DataTable
            Dim rstBL As New System.Data.DataTable
            ' Dim _program As CT.Data.Program
            Dim _Engine As CT.Data.Engine
            Dim _transmission As CT.Data.Transmission
            ' _program = New CT.Data.Program()
            sTnDRegion = Region  '_program.GetXccLead(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.HCID)
            _Engine = New CT.Data.Engine()
            rst = _Engine.GetXccEngineList(sTnDRegion)
            If Not rst Is Nothing Then
                For Each row As DataRow In rst.Rows
                    colEngineData.Add(row("FuelType") & "~" & row("EngineName"))
                Next
            End If
            _transmission = New CT.Data.Transmission()
            rst = _transmission.GetXCCTransmissions(sTnDRegion)
            If rst IsNot Nothing Then
                For Each row As DataRow In rst.Rows
                    colTransData.Add(row("pe16_TransType") & "~" & row("pe07_TransName"))
                Next
            End If
        End Sub
    End Class
End Namespace
