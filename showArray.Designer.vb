<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class showArray
#Disable Warning IDE1006 'Inherited code with variable names this suppresses: These words must begin with upper case characters
#Disable Warning IDE0054 'Inherited code with assignments (60) this suppresses: Use compound assignment x += 5 vs x = x + 5
#Disable Warning IDE0044 'Make field readonly
#Disable Warning BC40000 'VB compatibility
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(showArray))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._imgShowIcon_0 = New System.Windows.Forms.PictureBox()
        Me._lblShowTitle_0 = New System.Windows.Forms.Label()
        Me._imgShowColor_0 = New System.Windows.Forms.PictureBox()
        Me.imgShowColor = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(Me.components)
        Me.imgShowIcon = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(Me.components)
        Me.lblShowTitle = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        CType(Me._imgShowIcon_0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._imgShowColor_0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgShowColor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgShowIcon, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblShowTitle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_imgShowIcon_0
        '
        Me._imgShowIcon_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._imgShowIcon_0.Image = CType(resources.GetObject("_imgShowIcon_0.Image"), System.Drawing.Image)
        Me.imgShowIcon.SetIndex(Me._imgShowIcon_0, CType(0, Short))
        Me._imgShowIcon_0.Location = New System.Drawing.Point(104, 32)
        Me._imgShowIcon_0.Name = "_imgShowIcon_0"
        Me._imgShowIcon_0.Size = New System.Drawing.Size(17, 17)
        Me._imgShowIcon_0.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._imgShowIcon_0.TabIndex = 0
        Me._imgShowIcon_0.TabStop = False
        Me._imgShowIcon_0.Visible = False
        '
        '_lblShowTitle_0
        '
        Me._lblShowTitle_0.BackColor = System.Drawing.Color.Transparent
        Me._lblShowTitle_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblShowTitle_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblShowTitle_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblShowTitle.SetIndex(Me._lblShowTitle_0, CType(0, Short))
        Me._lblShowTitle_0.Location = New System.Drawing.Point(9, 32)
        Me._lblShowTitle_0.Name = "_lblShowTitle_0"
        Me._lblShowTitle_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblShowTitle_0.Size = New System.Drawing.Size(89, 17)
        Me._lblShowTitle_0.TabIndex = 0
        Me._lblShowTitle_0.Text = "8-Quantum Theory"
        Me._lblShowTitle_0.Visible = False
        '
        '_imgShowColor_0
        '
        Me._imgShowColor_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._imgShowColor_0.Image = CType(resources.GetObject("_imgShowColor_0.Image"), System.Drawing.Image)
        Me.imgShowColor.SetIndex(Me._imgShowColor_0, CType(0, Short))
        Me._imgShowColor_0.Location = New System.Drawing.Point(8, 32)
        Me._imgShowColor_0.Name = "_imgShowColor_0"
        Me._imgShowColor_0.Size = New System.Drawing.Size(113, 17)
        Me._imgShowColor_0.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._imgShowColor_0.TabIndex = 1
        Me._imgShowColor_0.TabStop = False
        Me._imgShowColor_0.Visible = False
        '
        'imgShowColor
        '
        '
        'imgShowIcon
        '
        '
        'lblShowTitle
        '
        '
        'showArray
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(289, 581)
        Me.ControlBox = True
        Me.Controls.Add(Me._imgShowIcon_0)
        Me.Controls.Add(Me._lblShowTitle_0)
        Me.Controls.Add(Me._imgShowColor_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(4, 30)
        Me.Name = "showArray"
        Me.Text = "Stack of Cards - Window needs to STay"
        Me.TopMost = True
        CType(Me._imgShowIcon_0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._imgShowColor_0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgShowColor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgShowIcon, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblShowTitle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class
