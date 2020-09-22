Public Class CtBaseClass

    Protected _tbAnswer As DataTable = Nothing
    Protected _arrayDT As String(,) = Nothing


    ''' <summary>
    ''' The method converts the DataTable which is returned from Database to string Array 
    ''' because only array string can be written in a Excel Range.
    ''' </summary>
    Protected Overridable Sub ConvertDataTableToStingArray()
        _arrayDT = Nothing
        Dim i, j As Integer
        If _tbAnswer IsNot Nothing Then

            ReDim _arrayDT(_tbAnswer.Rows.Count, _tbAnswer.Columns.Count)
            For i = 0 To _tbAnswer.Rows.Count - 1
                For j = 0 To _tbAnswer.Columns.Count - 1
                    _arrayDT(i, j) = _tbAnswer.Rows(i)(j).ToString()
                Next j
            Next i
        End If
    End Sub



    Protected Overridable Sub ConvertDataTableWithColumnToStingArray()
        _arrayDT = Nothing
        Dim i, j As Integer
        If _tbAnswer IsNot Nothing Then
            ReDim _arrayDT(_tbAnswer.Rows.Count, _tbAnswer.Columns.Count - 1)

            i = 0
            For j = 0 To _tbAnswer.Columns.Count - 1
                _arrayDT(i, j) = _tbAnswer.Columns(j).ColumnName.ToString
            Next j

            For i = 0 To _tbAnswer.Rows.Count - 1
                For j = 0 To _tbAnswer.Columns.Count - 1
                    _arrayDT(i + 1, j) = _tbAnswer.Rows(i)(j).ToString()
                Next j
            Next i
        End If
    End Sub


End Class
