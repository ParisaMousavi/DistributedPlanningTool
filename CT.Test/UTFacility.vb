Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UTFacility

    <TestMethod()> Public Sub DAT_GetCbg()

        Dim _GetCbg As CT.Data.Facility = New Data.Facility

        Assert.IsTrue(_GetCbg.GetCbg(Nothing, Nothing, Nothing, Nothing).Rows.Count > 0)

    End Sub

    <TestMethod()> Public Sub DAT_GetLocation()

        Dim _DAT_GetLocation As CT.Data.Facility = New Data.Facility

        Assert.IsTrue(_DAT_GetLocation.GetLocation("FoE", Nothing, Nothing, Nothing).Rows.Count > 0)

    End Sub

    <TestMethod()> Public Sub DAT_GetName()

        Dim _DAT_GetName As CT.Data.Facility = New Data.Facility

        Assert.IsTrue(_DAT_GetName.GetName("FoE", "MEC", Nothing, Nothing).Rows.Count > 0)

    End Sub

    <TestMethod()> Public Sub DAT_GetSubName()

        Dim _DAT_GetSubName As CT.Data.Facility = New Data.Facility

        Assert.IsTrue(_DAT_GetSubName.GetSubName("FoE", "MEC", "Windtunnel", Nothing).Rows.Count > 0)

    End Sub

End Class