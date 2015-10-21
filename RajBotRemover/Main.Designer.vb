<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btRemovePC = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btCleanDrives = New System.Windows.Forms.Button()
        Me.txtlog = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btRemovePC
        '
        Me.btRemovePC.Location = New System.Drawing.Point(12, 29)
        Me.btRemovePC.Name = "btRemovePC"
        Me.btRemovePC.Size = New System.Drawing.Size(392, 23)
        Me.btRemovePC.TabIndex = 0
        Me.btRemovePC.Text = "Remove Rajbot virus from this computer"
        Me.btRemovePC.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(191, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Label1"
        '
        'btCleanDrives
        '
        Me.btCleanDrives.Location = New System.Drawing.Point(13, 59)
        Me.btCleanDrives.Name = "btCleanDrives"
        Me.btCleanDrives.Size = New System.Drawing.Size(391, 23)
        Me.btCleanDrives.TabIndex = 2
        Me.btCleanDrives.Text = "Clean all usb drives on that computer"
        Me.btCleanDrives.UseVisualStyleBackColor = True
        '
        'txtlog
        '
        Me.txtlog.Location = New System.Drawing.Point(13, 93)
        Me.txtlog.Multiline = True
        Me.txtlog.Name = "txtlog"
        Me.txtlog.ReadOnly = True
        Me.txtlog.Size = New System.Drawing.Size(391, 239)
        Me.txtlog.TabIndex = 3
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(421, 344)
        Me.Controls.Add(Me.txtlog)
        Me.Controls.Add(Me.btCleanDrives)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btRemovePC)
        Me.Name = "Main"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "RajbotRemover"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btRemovePC As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btCleanDrives As System.Windows.Forms.Button
    Friend WithEvents txtlog As System.Windows.Forms.TextBox

End Class
