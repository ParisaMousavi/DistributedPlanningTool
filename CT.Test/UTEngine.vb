Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UTEngine

    <TestMethod()> Public Sub DAT_GetEngine()

        Dim _GetEngine As CT.Data.Engine = New Data.Engine

        Assert.IsTrue(_GetEngine.GetXccEngineList("Foe").Rows.Count > 0)

    End Sub


End Class