Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UTInstrumentation

    <TestMethod()> Public Sub DAT_Delete()

        Dim _Instrumentation As CT.Data.SevenTabs.Instrumentation = New Data.SevenTabs.Instrumentation()

        Assert.IsTrue(_Instrumentation.Delete(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "A new Column", "Pre-Instr. Parts", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_AddColumn()

        Dim _Instrumentation As CT.Data.SevenTabs.Instrumentation = New Data.SevenTabs.Instrumentation()

        Assert.IsTrue(_Instrumentation.AddColumn(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "A new Column", "Pre-Instr. Parts", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_EditColumn()

        Dim _Instrumentation As CT.Data.SevenTabs.Instrumentation = New Data.SevenTabs.Instrumentation()

        Assert.IsTrue(_Instrumentation.EditColumn(GlobalValues.Pe01, GlobalValues.Pe02, GlobalValues.HCID, "A new Column", "instrumented Dat Test AAA", "Pre-Instr. Parts", "Pre-Instr. Parts", GlobalValues.BuildType) = True)

    End Sub

    <TestMethod()> Public Sub DAT_UpdateData()

        Dim _Instrumentation As CT.Data.SevenTabs.Instrumentation = New Data.SevenTabs.Instrumentation()

        Assert.IsTrue(_Instrumentation.UpdateData(GlobalValues.Pe02, GlobalValues.pe45, "instrumented part 4", "Pre-Instr. Parts", "DAT Data Update", GlobalValues.BuildType) = True)

    End Sub

End Class
