Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UTMfcSpecification

    <TestMethod()> Public Sub DAT_Delete()

        Dim _MfcSpecification As CT.Data.SevenTabs.MfcSpecification = New Data.SevenTabs.MfcSpecification()

        Assert.IsTrue(_MfcSpecification.Delete(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "WA = Market", "Section 1 - Vehicle General", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_AddColumn()

        Dim _MfcSpecification As CT.Data.SevenTabs.MfcSpecification = New Data.SevenTabs.MfcSpecification()

        Assert.IsTrue(_MfcSpecification.AddColumn(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "MFC New Column", "Section 1 - Vehicle General", "New Column Description", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_EditColumn()

        Dim _MfcSpecification As CT.Data.SevenTabs.MfcSpecification = New Data.SevenTabs.MfcSpecification()

        Assert.IsTrue(_MfcSpecification.EditColumn(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "NC = New Column - Edit", "DAT EDIT", "Updated", "Section 1 - Vehicle General", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_UpdateData()

        Dim _MfcSpecification As CT.Data.SevenTabs.MfcSpecification = New Data.SevenTabs.MfcSpecification()

        Assert.IsTrue(_MfcSpecification.UpdateData(GlobalValues.Pe02, GlobalValues.pe45, "MFC New Column Edit", "Section 1 - Vehicle General", "DAT Value", GlobalValues.BuildType) = True)

    End Sub

End Class
