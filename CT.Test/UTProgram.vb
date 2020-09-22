Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UTProgram

    <TestMethod()> Public Sub DAT_GetXccLead()

        Dim _Program As CT.Data.Program = New Data.Program()
        Assert.IsTrue(_Program.GetXccLead(GlobalValues.Pe01, GlobalValues.HCID) = "FoE")

    End Sub

End Class