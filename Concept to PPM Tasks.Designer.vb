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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.ButBrowse = New System.Windows.Forms.Button()
        Me.Path = New System.Windows.Forms.TextBox()
        Me.ButSubmit = New System.Windows.Forms.Button()
        Me.Version = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'ButBrowse
        '
        Me.ButBrowse.Location = New System.Drawing.Point(263, 10)
        Me.ButBrowse.Name = "ButBrowse"
        Me.ButBrowse.Size = New System.Drawing.Size(75, 23)
        Me.ButBrowse.TabIndex = 0
        Me.ButBrowse.Text = "Browse"
        Me.ButBrowse.UseVisualStyleBackColor = True
        '
        'Path
        '
        Me.Path.Location = New System.Drawing.Point(29, 12)
        Me.Path.Name = "Path"
        Me.Path.ReadOnly = True
        Me.Path.Size = New System.Drawing.Size(228, 20)
        Me.Path.TabIndex = 1
        '
        'ButSubmit
        '
        Me.ButSubmit.Location = New System.Drawing.Point(263, 39)
        Me.ButSubmit.Name = "ButSubmit"
        Me.ButSubmit.Size = New System.Drawing.Size(75, 21)
        Me.ButSubmit.TabIndex = 2
        Me.ButSubmit.Text = "Submit"
        Me.ButSubmit.UseVisualStyleBackColor = True
        '
        'Version
        '
        Me.Version.AutoSize = True
        Me.Version.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Version.Location = New System.Drawing.Point(0, 64)
        Me.Version.Name = "Version"
        Me.Version.Size = New System.Drawing.Size(32, 12)
        Me.Version.TabIndex = 3
        Me.Version.Text = "Label1"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(122, 62)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Label1"
        Me.Label1.Visible = False
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"Concept", "AXIM"})
        Me.ComboBox1.Location = New System.Drawing.Point(29, 39)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(228, 21)
        Me.ComboBox1.TabIndex = 5
        Me.ComboBox1.Text = "Concept"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(346, 77)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Version)
        Me.Controls.Add(Me.ButSubmit)
        Me.Controls.Add(Me.Path)
        Me.Controls.Add(Me.ButBrowse)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.Text = "CSV Converter"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents ButBrowse As Button
    Friend WithEvents Path As TextBox
    Friend WithEvents ButSubmit As Button
    Friend WithEvents Version As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents ComboBox1 As ComboBox
End Class
