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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Path = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Browse = New System.Windows.Forms.Button()
        Me.Gen = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(396, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Tag Processor Generator"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Path
        '
        Me.Path.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Path.Location = New System.Drawing.Point(74, 70)
        Me.Path.Name = "Path"
        Me.Path.Size = New System.Drawing.Size(297, 20)
        Me.Path.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(42, 73)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(26, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "File:"
        '
        'Browse
        '
        Me.Browse.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Browse.Location = New System.Drawing.Point(370, 68)
        Me.Browse.Name = "Browse"
        Me.Browse.Size = New System.Drawing.Size(31, 23)
        Me.Browse.TabIndex = 3
        Me.Browse.Text = "..."
        Me.Browse.UseVisualStyleBackColor = True
        '
        'Gen
        '
        Me.Gen.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Gen.Location = New System.Drawing.Point(185, 101)
        Me.Gen.Name = "Gen"
        Me.Gen.Size = New System.Drawing.Size(75, 23)
        Me.Gen.TabIndex = 4
        Me.Gen.Text = "Generate"
        Me.Gen.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.Location = New System.Drawing.Point(15, 30)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(393, 35)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Takes the Excel file that contains the template definitions and creates the tag p" & _
    "rocessor map."
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Main
        '
        Me.AllowDrop = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(420, 136)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Gen)
        Me.Controls.Add(Me.Browse)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Path)
        Me.Controls.Add(Me.Label1)
        Me.MinimumSize = New System.Drawing.Size(436, 174)
        Me.Name = "Main"
        Me.Text = "TagProcGen"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Path As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Browse As System.Windows.Forms.Button
    Friend WithEvents Gen As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label3 As System.Windows.Forms.Label

End Class
