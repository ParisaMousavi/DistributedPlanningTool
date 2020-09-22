


Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UTUnit

    <TestMethod()> Public Sub DAT_ChangeBuildSeq()

        Dim _Unit As CT.Data.VehiclePlan.Unit = New Data.VehiclePlan.Unit()


    End Sub

    <TestMethod()> Public Sub DAT_AddUnit()

        Dim _Unit As CT.Data.VehiclePlan.Unit = New Data.VehiclePlan.Unit()

        Dim pe03 As Long = 0
        Dim pe45 As Long = 0
        Dim GenericSplitRowNumber As Integer = 0

        _Unit.AddUnit(GlobalValues.Pe01, GlobalValues.Pe02, "VP", "Vehicle", GlobalValues.HCID, pe03, pe45, GenericSplitRowNumber, GlobalValues.BuildType) 'task id 14

        Assert.IsTrue(pe03 <> 0)
        Assert.IsTrue(pe45 <> 0)

    End Sub


    <TestMethod()> Public Sub DAT_Deletevehicle()

        Dim _Unit As CT.Data.VehiclePlan.Unit = New Data.VehiclePlan.Unit()


    End Sub

End Class
