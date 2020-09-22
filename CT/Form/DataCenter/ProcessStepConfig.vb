
Namespace Form.DataCenter

    Friend NotInheritable Class ProcessStepConfig
        Public Shared ReadOnly Property ProcessStepPe26 As Long
            Get
                Try
                    ProcessStepPe26 = Val(Form.DataCenter.GlobalValues.WS.Application.Selection.cells(1, 1).Formula.ToString.Split(";")(0).Replace("=CellFace(", "").Replace("""", "").Trim())
                Catch
                    ProcessStepPe26 = 0
                End Try
            End Get
        End Property

        Public Shared ReadOnly Property ProcessStepAllocatedUsercase As Integer

            Get
                Try
                    ProcessStepAllocatedUsercase = Val(Form.DataCenter.GlobalValues.WS.Application.Selection.cells(1, 1).Formula.ToString.Split(";")(1))
                Catch
                    ProcessStepAllocatedUsercase = 0
                End Try

            End Get

        End Property

        Public Shared ReadOnly Property ProcessStepSequence As Integer

            Get
                Try
                    ProcessStepSequence = Val(Form.DataCenter.GlobalValues.WS.Application.Selection.cells(1, 1).Formula.ToString.Split(";")(2))
                Catch
                    ProcessStepSequence = 0
                End Try

            End Get

        End Property
        Public Shared ReadOnly Property ProcessStepUserCase As String

            Get
                Try
                    ProcessStepUserCase = Form.DataCenter.GlobalValues.WS.Application.Selection.cells(1, 1).Formula.ToString.Split(";")(3).ToString
                Catch
                    ProcessStepUserCase = ""
                End Try

            End Get

        End Property
        Public Shared ReadOnly Property PSIsGapOrDelay As Boolean

            Get
                Try
                    If Form.DataCenter.GlobalValues.WS.Application.Selection.cells(1, 1).Formula.ToString.Split(";")(4).ToString = "Gap" Or
                        Form.DataCenter.GlobalValues.WS.Application.Selection.cells(1, 1).Formula.ToString.Split(";")(4).ToString = "Delay" Then
                        PSIsGapOrDelay = True
                    Else
                        PSIsGapOrDelay = False
                    End If
                Catch ex As Exception
                    PSIsGapOrDelay = False
                End Try

            End Get

        End Property
        Public Shared ReadOnly Property ProcessStepStartDate As Date

            Get
                Try
                    ProcessStepStartDate = CDate(Form.DataCenter.GlobalValues.WS.Cells(4, Form.DataCenter.GlobalValues.WS.Application.Selection.cells(1, 1).Column).value2)
                Catch
                    ProcessStepStartDate = Nothing
                End Try

            End Get

        End Property
        Public Shared ReadOnly Property ProcessStepEndDate As Date

            Get
                Try
                    ProcessStepEndDate = CDate(Form.DataCenter.GlobalValues.WS.Cells(4, Form.DataCenter.GlobalValues.WS.Application.Selection.cells(1, Form.DataCenter.GlobalValues.WS.Application.Selection.cells.count + 1).Column).value2)
                Catch
                    ProcessStepEndDate = Nothing
                End Try

            End Get

        End Property


        'Private Function GetVehiclePS() As String

        '    GetVehiclePS = ""

        '    'Dim colPS As New Collection
        '    'Dim intcnt As Integer = 0, rngFind As Excel.Range = Nothing
        '    'Dim intFCol As Integer = 0, intLcol As Integer = 0, intRow As Integer = 0, intPe26 As Long = 0
        '    'Dim m_stAddress As String = "", intPS As Integer = 0, intUC As Integer = 0, dtStartDate As Date = Nothing

        '    'With Form.DataCenter.WS
        '    '    intRow = .Application.Selection.cells(1, 1).row
        '    '    Form.DisplayUtilities.Utilities.FindFLCols(intRow, intFCol, intLcol)
        '    '    With .Range(.Cells(intRow, intFCol - 1), .Cells(intRow, intLcol))
        '    '        rngFind = .Find("*", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
        '    '        If rngFind IsNot Nothing Then
        '    '            m_stAddress = rngFind.Address
        '    '            Do
        '    '                colPS.Add(rngFind.Formula & "~" & CDate(Form.DataCenter.WS.Cells(4, rngFind.Column).value2))
        '    '                rngFind = .FindNext(rngFind)
        '    '            Loop While Not rngFind Is Nothing And rngFind.Address <> m_stAddress
        '    '        End If
        '    '    End With
        '    '    Dim strSplit() As String = Nothing, strUC As String = "", strPS As String = ""
        '    '    For intcnt = 1 To colPS.Count
        '    '        If colPS.Item(intcnt).ToString.Split("~")(0) = "-" Then
        '    '            strUC = "-"
        '    '            strPS = "-"
        '    '            intPS = 0
        '    '        Else
        '    '            strSplit = colPS.Item(intcnt).ToString.Split("~")(0).Split(";")
        '    '            If strUC <> strSplit(3) Then
        '    '                strUC = strSplit(3)
        '    '                intUC = intUC + 1
        '    '                intPS = 0
        '    '            End If
        '    '            If strPS <> strSplit(0).Replace("=CellFace(", "").Replace("""", "").Trim() Then
        '    '                strPS = strSplit(0).Replace("=CellFace(", "").Replace("""", "").Trim()
        '    '                intPS = intPS + 1
        '    '            End If
        '    '            If .Application.Selection.cells(1, 1).Formula.ToString.Split(";")(0).Replace("=CellFace(", "").Replace("""", "").Trim() = strPS Then
        '    '                GetVehiclePS = strPS & "~" & intUC & "~" & intPS & "~" & colPS.Item(intcnt).ToString.Split("~")(1)
        '    '                Exit For
        '    '            End If
        '    '        End If
        '    '    Next
        '    'End With
        'End Function
    End Class
End Namespace