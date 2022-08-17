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
        Me.Find = New System.Windows.Forms.Label()
        Me.Replace = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.sFileType = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Find
        '
        Me.Find.AutoSize = True
        Me.Find.Location = New System.Drawing.Point(232, 43)
        Me.Find.Name = "Find"
        Me.Find.Size = New System.Drawing.Size(26, 12)
        Me.Find.TabIndex = 0
        Me.Find.Text = "Find"
        '
        'Replace
        '
        Me.Replace.AutoSize = True
        Me.Replace.Location = New System.Drawing.Point(232, 99)
        Me.Replace.Name = "Replace"
        Me.Replace.Size = New System.Drawing.Size(42, 12)
        Me.Replace.TabIndex = 1
        Me.Replace.Text = "Replace"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(28, 33)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(120, 57)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Replace"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(299, 33)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(137, 22)
        Me.TextBox1.TabIndex = 3
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(299, 89)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(137, 22)
        Me.TextBox2.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(20, 150)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(58, 12)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Folder path"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(84, 147)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(352, 22)
        Me.TextBox3.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(100, 245)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(18, 12)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Ty"
        '
        'sFileType
        '
        Me.sFileType.Location = New System.Drawing.Point(109, 192)
        Me.sFileType.Name = "sFileType"
        Me.sFileType.Size = New System.Drawing.Size(100, 22)
        Me.sFileType.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(20, 202)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(83, 12)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Replace file type"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(509, 306)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.sFileType)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Replace)
        Me.Controls.Add(Me.Find)
        Me.Name = "Form1"
        Me.Text = "File Contact Replace"
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub
    Friend WithEvents Find As System.Windows.Forms.Label
    Friend WithEvents Replace As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents sFileType As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label

End Class
