Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UTFurtherBasicSpecification

    <TestMethod()> Public Sub DAT_Delete()

        Dim _FurtherBasicSpecification As CT.Data.SevenTabs.FurtherBasicSpecification = New Data.SevenTabs.FurtherBasicSpecification()

        Assert.IsTrue(_FurtherBasicSpecification.Delete(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "Start/Stop", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_AddColumn()

        Dim _FurtherBasicSpecification As CT.Data.SevenTabs.FurtherBasicSpecification = New Data.SevenTabs.FurtherBasicSpecification()

        Assert.IsTrue(_FurtherBasicSpecification.AddColumn(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "DAT Structure Test", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_EditColumn()

        Dim _FurtherBasicSpecification As CT.Data.SevenTabs.FurtherBasicSpecification = New Data.SevenTabs.FurtherBasicSpecification()

        Assert.IsTrue(_FurtherBasicSpecification.EditColumn(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "DAT Structure Test Updated", "DAT Structure test Updated II", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_UpdateData()

        Dim _FurtherBasicSpecification As CT.Data.SevenTabs.FurtherBasicSpecification = New Data.SevenTabs.FurtherBasicSpecification()

        Assert.IsTrue(_FurtherBasicSpecification.UpdateData(GlobalValues.Pe02, GlobalValues.pe45, "DAT Structure test Updated II", "DAT Value III", GlobalValues.BuildType) = True)

    End Sub

End Class