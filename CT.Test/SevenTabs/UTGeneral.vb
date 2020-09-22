Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UTGeneral

    <TestMethod()> Public Sub DAT_GetDynamicHeaders()

        Dim _General As CT.Data.SevenTabs.General = New Data.SevenTabs.General
        Assert.IsTrue(_General.GetDynamicHeaders(GlobalValues.Pe01, GlobalValues.HCID, GlobalValues.BuildPhase).Rows.Count > 0)



    End Sub

End Class