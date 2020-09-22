Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UTUtilities

    <TestMethod()> Public Sub DAT_GetChangelogData()

        Dim _GetChangelogData As CT.Data.Utilities = New Data.Utilities()

        Assert.IsTrue(_GetChangelogData.GetChangelogdata(GlobalValues.Pe02).Rows.Count > 0)

    End Sub

End Class
