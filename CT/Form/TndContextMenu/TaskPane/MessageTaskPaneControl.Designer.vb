<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MessageTaskPaneControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MessageTaskPaneControl))
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TSMessage = New System.Windows.Forms.ToolStrip()
        Me.btnRefresh = New System.Windows.Forms.ToolStripButton()
        Me.btnMarkAsRead = New System.Windows.Forms.ToolStripButton()
        Me.tsHide = New System.Windows.Forms.ToolStripButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtMessage = New System.Windows.Forms.TextBox()
        Me.lstMessages = New System.Windows.Forms.ListView()
        Me.CreatedBy = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.CreatedOn = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.MessageText = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.pe94_TnDPlanMessage_PK = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.TSMessage.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(3, 35)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 13)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Inbox"
        '
        'TSMessage
        '
        Me.TSMessage.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.TSMessage.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnRefresh, Me.btnMarkAsRead, Me.tsHide})
        Me.TSMessage.Location = New System.Drawing.Point(0, 0)
        Me.TSMessage.Name = "TSMessage"
        Me.TSMessage.Size = New System.Drawing.Size(290, 27)
        Me.TSMessage.TabIndex = 8
        Me.TSMessage.Text = "Messages"
        '
        'btnRefresh
        '
        Me.btnRefresh.Image = CType(resources.GetObject("btnRefresh.Image"), System.Drawing.Image)
        Me.btnRefresh.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(70, 24)
        Me.btnRefresh.Text = "Refresh"
        '
        'btnMarkAsRead
        '
        Me.btnMarkAsRead.Image = CType(resources.GetObject("btnMarkAsRead.Image"), System.Drawing.Image)
        Me.btnMarkAsRead.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnMarkAsRead.Name = "btnMarkAsRead"
        Me.btnMarkAsRead.Size = New System.Drawing.Size(98, 24)
        Me.btnMarkAsRead.Text = "Mark as read"
        '
        'tsHide
        '
        Me.tsHide.Image = CType(resources.GetObject("tsHide.Image"), System.Drawing.Image)
        Me.tsHide.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsHide.Name = "tsHide"
        Me.tsHide.Size = New System.Drawing.Size(56, 24)
        Me.tsHide.Text = "Hide"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(-3, 553)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(82, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Message text"
        '
        'txtMessage
        '
        Me.txtMessage.Location = New System.Drawing.Point(0, 578)
        Me.txtMessage.Multiline = True
        Me.txtMessage.Name = "txtMessage"
        Me.txtMessage.ReadOnly = True
        Me.txtMessage.Size = New System.Drawing.Size(273, 59)
        Me.txtMessage.TabIndex = 6
        '
        'lstMessages
        '
        Me.lstMessages.CheckBoxes = True
        Me.lstMessages.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.CreatedBy, Me.CreatedOn, Me.MessageText, Me.pe94_TnDPlanMessage_PK})
        Me.lstMessages.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lstMessages.Location = New System.Drawing.Point(0, 65)
        Me.lstMessages.MultiSelect = False
        Me.lstMessages.Name = "lstMessages"
        Me.lstMessages.Size = New System.Drawing.Size(273, 485)
        Me.lstMessages.TabIndex = 5
        Me.lstMessages.UseCompatibleStateImageBehavior = False
        Me.lstMessages.View = System.Windows.Forms.View.Details
        '
        'CreatedBy
        '
        Me.CreatedBy.Text = "Updated by"
        Me.CreatedBy.Width = 75
        '
        'CreatedOn
        '
        Me.CreatedOn.Text = "Date/Time"
        Me.CreatedOn.Width = 90
        '
        'MessageText
        '
        Me.MessageText.Text = "Message Text"
        Me.MessageText.Width = 150
        '
        'pe94_TnDPlanMessage_PK
        '
        Me.pe94_TnDPlanMessage_PK.Text = "pe94_TnDPlanMessage_PK"
        Me.pe94_TnDPlanMessage_PK.Width = 0
        '
        'MessageTaskPaneControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TSMessage)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtMessage)
        Me.Controls.Add(Me.lstMessages)
        Me.Name = "MessageTaskPaneControl"
        Me.Size = New System.Drawing.Size(290, 652)
        Me.TSMessage.ResumeLayout(False)
        Me.TSMessage.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TSMessage As System.Windows.Forms.ToolStrip
    Friend WithEvents btnRefresh As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnMarkAsRead As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsHide As System.Windows.Forms.ToolStripButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtMessage As System.Windows.Forms.TextBox
    Friend WithEvents lstMessages As System.Windows.Forms.ListView
    Friend WithEvents CreatedBy As System.Windows.Forms.ColumnHeader
    Friend WithEvents CreatedOn As System.Windows.Forms.ColumnHeader
    Friend WithEvents MessageText As System.Windows.Forms.ColumnHeader
    Friend WithEvents pe94_TnDPlanMessage_PK As System.Windows.Forms.ColumnHeader
End Class
