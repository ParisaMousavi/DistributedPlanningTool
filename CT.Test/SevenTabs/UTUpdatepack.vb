Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UTUpdatepack

    <TestMethod()> Public Sub DAT_Delete()

        Dim _Updatepack As CT.Data.SevenTabs.Updatepack = New Data.SevenTabs.Updatepack()

        Assert.IsTrue(_Updatepack.Delete(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "Update#1", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_AddColumn()

        Dim _Updatepack As CT.Data.SevenTabs.Updatepack = New Data.SevenTabs.Updatepack()

        Assert.IsTrue(_Updatepack.AddColumn(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "DAT UP 1", "Desc", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_EditColumn()

        Dim _Updatepack As CT.Data.SevenTabs.Updatepack = New Data.SevenTabs.Updatepack()

        Assert.IsTrue(_Updatepack.EditColumn(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "DAT UP ' 1", "DAT UP 1", "Desc Edit", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_UpdateData()

        Dim _Updatepack As CT.Data.SevenTabs.Updatepack = New Data.SevenTabs.Updatepack()

        Assert.IsTrue(_Updatepack.UpdateData(GlobalValues.Pe02, GlobalValues.pe45, "DAT UP 1", "DAT Value", GlobalValues.BuildType) = True)

    End Sub

End Class
