<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.GroupBox10 = New System.Windows.Forms.GroupBox()
        Me.txtOTHmemo = New System.Windows.Forms.TextBox()
        Me.GroupBox10.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox10
        '
        Me.GroupBox10.Controls.Add(Me.txtOTHmemo)
        Me.GroupBox10.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox10.Name = "GroupBox10"
        Me.GroupBox10.Size = New System.Drawing.Size(455, 192)
        Me.GroupBox10.TabIndex = 51
        Me.GroupBox10.TabStop = False
        Me.GroupBox10.Text = "Краткое описание"
        '
        'txtOTHmemo
        '
        Me.txtOTHmemo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtOTHmemo.Location = New System.Drawing.Point(3, 16)
        Me.txtOTHmemo.Multiline = True
        Me.txtOTHmemo.Name = "txtOTHmemo"
        Me.txtOTHmemo.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtOTHmemo.Size = New System.Drawing.Size(449, 173)
        Me.txtOTHmemo.TabIndex = 54
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1078, 686)
        Me.Controls.Add(Me.GroupBox10)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.GroupBox10.ResumeLayout(False)
        Me.GroupBox10.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox10 As System.Windows.Forms.GroupBox
    Friend WithEvents txtOTHmemo As System.Windows.Forms.TextBox

End Class
