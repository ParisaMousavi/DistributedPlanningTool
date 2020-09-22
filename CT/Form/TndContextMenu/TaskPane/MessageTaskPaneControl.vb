Imports System.Windows.Forms
Public Class MessageTaskPaneControl
    Dim bolIsLoading As Boolean = False

    Public Sub loadData()
        Try

            If Me.Visible = False Then Exit Sub
            Me.Cursor = Cursors.WaitCursor
            LoadMessages()
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.MessagesTaskPane, ex.Message), "Show messages", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        LoadMessages()
    End Sub
    Public Sub LoadMessages()
        Try
            bolIsLoading = True
            Dim objDat As New CT.Data.MessagePassing
            Dim DT As System.Data.DataTable = objDat.SelectAll(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)
            Dim DR As System.Data.DataRow = Nothing
            lstMessages.Items.Clear()
            Dim li As ListViewItem = Nothing
            Dim liSub As ListViewItem.ListViewSubItem = Nothing
            Form.DataCenter.GlobalValues.CurrentTotalMessages = DT.Rows.Count
            For Each DR In DT.Rows
                li = lstMessages.Items.Add(DR("CDSID").ToString)
                liSub = New ListViewItem.ListViewSubItem
                liSub.Text = DR("InsertTime").ToString
                li.SubItems.Add(liSub)
                liSub = New ListViewItem.ListViewSubItem
                liSub.Text = DR("MessageText").ToString
                li.SubItems.Add(liSub)
                liSub = New ListViewItem.ListViewSubItem
                liSub.Text = DR("pe94_TnDPlanMessage_PK").ToString
                li.SubItems.Add(liSub)
            Next
            btnMarkAsRead.Enabled = False
            txtMessage.Text = String.Empty

        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.MessagesTaskPane, ex.Message), "Show messages", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            bolIsLoading = False
        End Try
    End Sub
    Private Sub lstMessages_Click(sender As Object, e As EventArgs) Handles lstMessages.Click
        ShowMessage()
    End Sub
    Private Sub lstMessages_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) Handles lstMessages.ItemSelectionChanged
        ShowMessage()
    End Sub
    Private Sub ShowMessage()
        Try
            If bolIsLoading = True Then Exit Sub
            If lstMessages.CheckedItems.Count > 0 Then
                btnMarkAsRead.Enabled = True
            Else
                btnMarkAsRead.Enabled = False
            End If
            txtMessage.Text = String.Empty
            Try
                txtMessage.Text = lstMessages.SelectedItems(0).SubItems(2).Text
            Catch ex As Exception
            End Try
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.MessagesTaskPane, ex.Message), "Show messages", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnMarkAsRead_Click(sender As Object, e As EventArgs) Handles btnMarkAsRead.Click
        Try
            Dim objDat As New CT.Data.MessagePassing
            Dim li As ListViewItem
            For Each li In lstMessages.CheckedItems
                If objDat.SetAsRead(CInt(Val(li.SubItems(3).Text))) = False Then
                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                End If
            Next
            LoadMessages()
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.MessagesTaskPane, ex.Message), "Show messages", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub lstMessages_ItemActivate(sender As Object, e As EventArgs) Handles lstMessages.ItemActivate
        ShowMessage()
    End Sub
    Private Sub lstMessages_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles lstMessages.ItemChecked
        ShowMessage()
    End Sub

    Private Sub tsHide_Click(sender As Object, e As EventArgs) Handles tsHide.Click

        Try
            Globals.Ribbons.RbnTnDControlPanel.TGMessages.Checked = False
            Form.DataCenter.GlobalValues.wsEve.ShowMessageTaskPane(False)
        Catch ex As Exception
        End Try

    End Sub

    Private Sub lstMessages_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstMessages.SelectedIndexChanged
        ShowMessage()
    End Sub

End Class
