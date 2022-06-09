<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class showArray
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents _imgShowIcon_0 As System.Windows.Forms.PictureBox
	Public WithEvents _lblShowTitle_0 As System.Windows.Forms.Label
	Public WithEvents _imgShowColor_0 As System.Windows.Forms.PictureBox
	Public WithEvents imgShowColor As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
	Public WithEvents imgShowIcon As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
	Public WithEvents lblShowTitle As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(showArray))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me._imgShowIcon_0 = New System.Windows.Forms.PictureBox
		Me._lblShowTitle_0 = New System.Windows.Forms.Label
		Me._imgShowColor_0 = New System.Windows.Forms.PictureBox
		Me.imgShowColor = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
		Me.imgShowIcon = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
		Me.lblShowTitle = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.imgShowColor, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.imgShowIcon, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblShowTitle, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "Stack of Cards"
		Me.ClientSize = New System.Drawing.Size(289, 581)
		Me.Location = New System.Drawing.Point(4, 30)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "showArray"
		Me._imgShowIcon_0.Size = New System.Drawing.Size(17, 17)
		Me._imgShowIcon_0.Location = New System.Drawing.Point(104, 32)
		Me._imgShowIcon_0.Image = CType(resources.GetObject("_imgShowIcon_0.Image"), System.Drawing.Image)
		Me._imgShowIcon_0.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
		Me._imgShowIcon_0.Visible = False
		Me._imgShowIcon_0.Enabled = True
		Me._imgShowIcon_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._imgShowIcon_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._imgShowIcon_0.Name = "_imgShowIcon_0"
		Me._lblShowTitle_0.Text = "8-Quantum Theory"
		Me._lblShowTitle_0.Size = New System.Drawing.Size(89, 17)
		Me._lblShowTitle_0.Location = New System.Drawing.Point(9, 32)
		Me._lblShowTitle_0.TabIndex = 0
		Me._lblShowTitle_0.Visible = False
		Me._lblShowTitle_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblShowTitle_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblShowTitle_0.BackColor = System.Drawing.Color.Transparent
		Me._lblShowTitle_0.Enabled = True
		Me._lblShowTitle_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblShowTitle_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblShowTitle_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblShowTitle_0.UseMnemonic = True
		Me._lblShowTitle_0.AutoSize = False
		Me._lblShowTitle_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblShowTitle_0.Name = "_lblShowTitle_0"
		Me._imgShowColor_0.Size = New System.Drawing.Size(113, 17)
		Me._imgShowColor_0.Location = New System.Drawing.Point(8, 32)
		Me._imgShowColor_0.Image = CType(resources.GetObject("_imgShowColor_0.Image"), System.Drawing.Image)
		Me._imgShowColor_0.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
		Me._imgShowColor_0.Visible = False
		Me._imgShowColor_0.Enabled = True
		Me._imgShowColor_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._imgShowColor_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._imgShowColor_0.Name = "_imgShowColor_0"
		Me.Controls.Add(_imgShowIcon_0)
		Me.Controls.Add(_lblShowTitle_0)
		Me.Controls.Add(_imgShowColor_0)
		Me.imgShowColor.SetIndex(_imgShowColor_0, CType(0, Short))
		Me.imgShowIcon.SetIndex(_imgShowIcon_0, CType(0, Short))
		Me.lblShowTitle.SetIndex(_lblShowTitle_0, CType(0, Short))
		CType(Me.lblShowTitle, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.imgShowIcon, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.imgShowColor, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class