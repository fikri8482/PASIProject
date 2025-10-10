<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSendMailExport
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSendMailExport))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtMsg = New System.Windows.Forms.TextBox()
        Me.grpMsg = New System.Windows.Forms.GroupBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rtbProcess = New System.Windows.Forms.RichTextBox()
        Me.txtSaveAsDOM = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtLast = New System.Windows.Forms.TextBox()
        Me.txtNext = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtSechedule = New System.Windows.Forms.TextBox()
        Me.txtAttachmentDOM = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnManual = New System.Windows.Forms.Button()
        Me.timerProcess = New System.Windows.Forms.Timer(Me.components)
        Me.lblDB = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.grpMsg.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Location = New System.Drawing.Point(-4, -2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(584, 64)
        Me.Panel1.TabIndex = 36
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(477, 43)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(83, 13)
        Me.Label12.TabIndex = 3
        Me.Label12.Text = "Version 1.0.0"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(109, 38)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(322, 18)
        Me.Label11.TabIndex = 2
        Me.Label11.Text = "SEND EXCEL TO SUPPLIER EXPORT"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Verdana", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(107, 7)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(434, 25)
        Me.Label10.TabIndex = 1
        Me.Label10.Text = "PT. AUTOCOMP SYSTEM INDONESIA"
        '
        'Panel2
        '
        Me.Panel2.BackgroundImage = CType(resources.GetObject("Panel2.BackgroundImage"), System.Drawing.Image)
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel2.Location = New System.Drawing.Point(4, 3)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(99, 57)
        Me.Panel2.TabIndex = 0
        '
        'txtMsg
        '
        Me.txtMsg.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMsg.BackColor = System.Drawing.Color.LightSteelBlue
        Me.txtMsg.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtMsg.Font = New System.Drawing.Font("Verdana", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMsg.ForeColor = System.Drawing.Color.Red
        Me.txtMsg.Location = New System.Drawing.Point(6, 13)
        Me.txtMsg.Multiline = True
        Me.txtMsg.Name = "txtMsg"
        Me.txtMsg.ReadOnly = True
        Me.txtMsg.Size = New System.Drawing.Size(554, 28)
        Me.txtMsg.TabIndex = 0
        Me.txtMsg.TabStop = False
        Me.txtMsg.Text = "Message"
        Me.txtMsg.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'grpMsg
        '
        Me.grpMsg.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpMsg.Controls.Add(Me.txtMsg)
        Me.grpMsg.Location = New System.Drawing.Point(5, 452)
        Me.grpMsg.Name = "grpMsg"
        Me.grpMsg.Size = New System.Drawing.Size(566, 44)
        Me.grpMsg.TabIndex = 37
        Me.grpMsg.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.rtbProcess)
        Me.GroupBox1.Controls.Add(Me.txtSaveAsDOM)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtLast)
        Me.GroupBox1.Controls.Add(Me.txtNext)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtSechedule)
        Me.GroupBox1.Controls.Add(Me.txtAttachmentDOM)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Location = New System.Drawing.Point(5, 69)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(566, 384)
        Me.GroupBox1.TabIndex = 42
        Me.GroupBox1.TabStop = False
        '
        'rtbProcess
        '
        Me.rtbProcess.Location = New System.Drawing.Point(20, 94)
        Me.rtbProcess.Name = "rtbProcess"
        Me.rtbProcess.Size = New System.Drawing.Size(534, 227)
        Me.rtbProcess.TabIndex = 59
        Me.rtbProcess.Text = ""
        '
        'txtSaveAsDOM
        '
        Me.txtSaveAsDOM.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSaveAsDOM.Location = New System.Drawing.Point(193, 40)
        Me.txtSaveAsDOM.Name = "txtSaveAsDOM"
        Me.txtSaveAsDOM.Size = New System.Drawing.Size(337, 21)
        Me.txtSaveAsDOM.TabIndex = 56
        Me.txtSaveAsDOM.Text = "D:\PASISystem\PASISystem\Template"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(17, 43)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 13)
        Me.Label1.TabIndex = 57
        Me.Label1.Text = "Result Folder"
        '
        'txtLast
        '
        Me.txtLast.BackColor = System.Drawing.SystemColors.Control
        Me.txtLast.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtLast.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLast.Location = New System.Drawing.Point(197, 327)
        Me.txtLast.Name = "txtLast"
        Me.txtLast.ReadOnly = True
        Me.txtLast.Size = New System.Drawing.Size(125, 21)
        Me.txtLast.TabIndex = 53
        Me.txtLast.TabStop = False
        '
        'txtNext
        '
        Me.txtNext.BackColor = System.Drawing.SystemColors.Control
        Me.txtNext.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNext.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNext.Location = New System.Drawing.Point(197, 354)
        Me.txtNext.Name = "txtNext"
        Me.txtNext.ReadOnly = True
        Me.txtNext.Size = New System.Drawing.Size(125, 21)
        Me.txtNext.TabIndex = 54
        Me.txtNext.TabStop = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(21, 356)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(90, 13)
        Me.Label8.TabIndex = 52
        Me.Label8.Text = "Next Process :"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(21, 329)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(87, 13)
        Me.Label7.TabIndex = 51
        Me.Label7.Text = "Last Process :"
        '
        'txtSechedule
        '
        Me.txtSechedule.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSechedule.Location = New System.Drawing.Point(193, 67)
        Me.txtSechedule.Name = "txtSechedule"
        Me.txtSechedule.Size = New System.Drawing.Size(70, 21)
        Me.txtSechedule.TabIndex = 6
        '
        'txtAttachmentDOM
        '
        Me.txtAttachmentDOM.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAttachmentDOM.Location = New System.Drawing.Point(193, 13)
        Me.txtAttachmentDOM.Name = "txtAttachmentDOM"
        Me.txtAttachmentDOM.Size = New System.Drawing.Size(337, 21)
        Me.txtAttachmentDOM.TabIndex = 5
        Me.txtAttachmentDOM.Text = "D:\TRIE\PASISystem\PASISystem\Template"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(17, 70)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(132, 13)
        Me.Label6.TabIndex = 41
        Me.Label6.Text = "Schedule Every (Seconds)"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(17, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(84, 13)
        Me.Label5.TabIndex = 40
        Me.Label5.Text = "Template Folder"
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.BackColor = System.Drawing.SystemColors.Control
        Me.btnExit.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.Image = Global.New_Send_To_Supplier_Export.My.Resources.Resources.door_out
        Me.btnExit.Location = New System.Drawing.Point(5, 507)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(87, 28)
        Me.btnExit.TabIndex = 10
        Me.btnExit.Tag = ""
        Me.btnExit.Text = "   &Exit"
        Me.btnExit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'btnManual
        '
        Me.btnManual.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnManual.BackColor = System.Drawing.SystemColors.Control
        Me.btnManual.Image = Global.New_Send_To_Supplier_Export.My.Resources.Resources.gear_in
        Me.btnManual.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnManual.Location = New System.Drawing.Point(459, 507)
        Me.btnManual.Name = "btnManual"
        Me.btnManual.Size = New System.Drawing.Size(112, 29)
        Me.btnManual.TabIndex = 7
        Me.btnManual.Text = "      &Manual Process"
        Me.btnManual.UseVisualStyleBackColor = False
        '
        'timerProcess
        '
        Me.timerProcess.Interval = 1000
        '
        'lblDB
        '
        Me.lblDB.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblDB.Font = New System.Drawing.Font("Verdana", 6.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDB.ForeColor = System.Drawing.Color.Black
        Me.lblDB.Location = New System.Drawing.Point(71, 544)
        Me.lblDB.Name = "lblDB"
        Me.lblDB.Size = New System.Drawing.Size(500, 13)
        Me.lblDB.TabIndex = 53
        Me.lblDB.Text = "Next Process : asdadasdasdsadas"
        Me.lblDB.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'frmSendMailExport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSteelBlue
        Me.ClientSize = New System.Drawing.Size(576, 562)
        Me.Controls.Add(Me.lblDB)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.grpMsg)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.btnManual)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSendMailExport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Tag = "SendToSupplier"
        Me.Text = "SEND EXCEL TO SUPPLIER EXPORT"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.grpMsg.ResumeLayout(False)
        Me.grpMsg.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnManual As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Public WithEvents txtMsg As System.Windows.Forms.TextBox
    Public WithEvents grpMsg As System.Windows.Forms.GroupBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtLast As System.Windows.Forms.TextBox
    Friend WithEvents txtNext As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtSechedule As System.Windows.Forms.TextBox
    Friend WithEvents txtAttachmentDOM As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtSaveAsDOM As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents timerProcess As System.Windows.Forms.Timer
    Friend WithEvents rtbProcess As System.Windows.Forms.RichTextBox
    Friend WithEvents lblDB As System.Windows.Forms.Label

End Class
