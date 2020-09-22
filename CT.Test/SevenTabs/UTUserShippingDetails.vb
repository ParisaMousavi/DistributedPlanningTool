Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UTUserShippingDetails

    <TestMethod()> Public Sub DAT_Delete()

        Dim _UserShippingDetails As CT.Data.SevenTabs.UserShippingDetails = New Data.SevenTabs.UserShippingDetails()

        Assert.IsTrue(_UserShippingDetails.Delete(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "CHP contact", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_AddColumn()

        Dim _UserShippingDetails As CT.Data.SevenTabs.UserShippingDetails = New Data.SevenTabs.UserShippingDetails()

        Assert.IsTrue(_UserShippingDetails.AddColumn(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "New DAT Column", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_EditColumn()

        Dim _UserShippingDetails As CT.Data.SevenTabsManagement.UserShippingDetails = New Data.SevenTabs.UserShippingDetails()

        Assert.IsTrue(_UserShippingDetails.EditColumn(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "New DAT Column", "DAT Edit Column", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_UpdateData()

        Dim _UserShippingDetails As CT.Data.SevenTabs.UserShippingDetails = New Data.SevenTabs.UserShippingDetails()

        Assert.IsTrue(_UserShippingDetails.UpdateData(GlobalValues.Pe02, GlobalValues.pe45, "DAT Edit Column", "DAT Value", GlobalValues.BuildType) = True)

    End Sub




End Class