Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UTNonMfcSpecification

    <TestMethod()> Public Sub DAT_Delete()

        Dim _NonMfcSpecification As CT.Data.SevenTabs.NonMfcSpecification = New Data.SevenTabs.NonMfcSpecification()

        Assert.IsTrue(_NonMfcSpecification.Delete(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "3000 Km run in", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_AddColumn()

        Dim _NonMfcSpecification As CT.Data.SevenTabs.NonMfcSpecification = New Data.SevenTabs.NonMfcSpecification()

        Assert.IsTrue(_NonMfcSpecification.AddColumn(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "DAT New NonMfc Column", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_EditColumn()

        Dim _NonMfcSpecification As CT.Data.SevenTabs.NonMfcSpecification = New Data.SevenTabs.NonMfcSpecification()

        Assert.IsTrue(_NonMfcSpecification.EditColumn(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "NC = New Column - Updated", "NC = New Column - Updated II", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_UpdateData()

        Dim _NonMfcSpecification As CT.Data.SevenTabs.NonMfcSpecification = New Data.SevenTabs.NonMfcSpecification()

        Assert.IsTrue(_NonMfcSpecification.UpdateData(GlobalValues.Pe02, GlobalValues.pe45, "Roll over cage New", "DAT Value", GlobalValues.BuildType) = True)

    End Sub

End Class
