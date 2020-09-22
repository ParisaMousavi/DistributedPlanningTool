Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UTProgramInformation

    <TestMethod()> Public Sub DAT_Delete()

        Dim _ProgramInformation As CT.Data.SevenTabs.ProgramInformation = New Data.SevenTabs.ProgramInformation()

        Assert.IsTrue(_ProgramInformation.Delete(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "Batch / Wave", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_AddColumn()

        Dim _ProgramInformation As CT.Data.SevenTabs.ProgramInformation = New Data.SevenTabs.ProgramInformation()

        Assert.IsTrue(_ProgramInformation.AddColumn(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "DAT New PI Column", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_EditColumn()

        Dim _ProgramInformation As CT.Data.SevenTabs.ProgramInformation = New Data.SevenTabs.ProgramInformation()

        Assert.IsTrue(_ProgramInformation.EditColumn(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "DAT Edi PI Column", "DAT Edit PI Column", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_UpdateData()

        Dim _ProgramInformation As CT.Data.SevenTabs.ProgramInformation = New Data.SevenTabs.ProgramInformation()

        Assert.IsTrue(_ProgramInformation.UpdateData(GlobalValues.Pe02, GlobalValues.pe45, "DAT Edi PI Column", "DAT Value", GlobalValues.BuildType) = True)

    End Sub

End Class
