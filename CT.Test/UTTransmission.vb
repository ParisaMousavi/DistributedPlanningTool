Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UTTransmission

    <TestMethod()> Public Sub DAT_GetTransmission()

        Dim _GetTransmission As CT.Data.Transmission = New Data.Transmission()

        Assert.IsTrue(_GetTransmission.GetXCCTransmissions("FoE").Rows.Count > 0)

    End Sub

End Class