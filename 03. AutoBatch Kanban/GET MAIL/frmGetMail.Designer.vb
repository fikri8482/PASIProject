<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGetMail
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmGetMail))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtMsg = New System.Windows.Forms.TextBox()
        Me.grpMsg = New System.Windows.Forms.GroupBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtPortE = New System.Windows.Forms.TextBox()
        Me.txtAttachmentE = New System.Windows.Forms.TextBox()
        Me.txtpop3E = New System.Windows.Forms.TextBox()
        Me.txtPasswordE = New System.Windows.Forms.TextBox()
        Me.txtUserNameE = New System.Windows.Forms.TextBox()
        Me.txtEmailAddressE = New System.Windows.Forms.TextBox()
        Me.rtbProcess = New System.Windows.Forms.RichTextBox()
        Me.txtLast = New System.Windows.Forms.TextBox()
        Me.txtNext = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtPort = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtSechedule = New System.Windows.Forms.TextBox()
        Me.txtAttachment = New System.Windows.Forms.TextBox()
        Me.txtpop3 = New System.Windows.Forms.TextBox()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.txtUserName = New System.Windows.Forms.TextBox()
        Me.txtEmailAddress = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtcounter = New System.Windows.Forms.TextBox()
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
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Location = New System.Drawing.Point(-4, -2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(797, 64)
        Me.Panel1.TabIndex = 36
        '
        'Label12
        '
        Me.Label12.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(691, 43)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(83, 13)
        Me.Label12.TabIndex = 3
        Me.Label12.Text = "Version 1.0.2"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(153, 38)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(78, 18)
        Me.Label11.TabIndex = 2
        Me.Label11.Text = "Get Mail"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Verdana", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(149, 7)
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
        Me.Panel2.Size = New System.Drawing.Size(119, 57)
        Me.Panel2.TabIndex = 0
        '
        'txtMsg
        '
        Me.txtMsg.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMsg.BackColor = System.Drawing.Color.LightSteelBlue
        Me.txtMsg.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtMsg.Font = New System.Drawing.Font("Verdana", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMsg.ForeColor = System.Drawing.Color.Red
        Me.txtMsg.Location = New System.Drawing.Point(6, 13)
        Me.txtMsg.Multiline = True
        Me.txtMsg.Name = "txtMsg"
        Me.txtMsg.ReadOnly = True
        Me.txtMsg.Size = New System.Drawing.Size(768, 28)
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
        Me.grpMsg.Location = New System.Drawing.Point(5, 469)
        Me.grpMsg.Name = "grpMsg"
        Me.grpMsg.Size = New System.Drawing.Size(780, 44)
        Me.grpMsg.TabIndex = 37
        Me.grpMsg.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.txtPortE)
        Me.GroupBox1.Controls.Add(Me.txtAttachmentE)
        Me.GroupBox1.Controls.Add(Me.txtpop3E)
        Me.GroupBox1.Controls.Add(Me.txtPasswordE)
        Me.GroupBox1.Controls.Add(Me.txtUserNameE)
        Me.GroupBox1.Controls.Add(Me.txtEmailAddressE)
        Me.GroupBox1.Controls.Add(Me.rtbProcess)
        Me.GroupBox1.Controls.Add(Me.txtLast)
        Me.GroupBox1.Controls.Add(Me.txtNext)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtPort)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.txtSechedule)
        Me.GroupBox1.Controls.Add(Me.txtAttachment)
        Me.GroupBox1.Controls.Add(Me.txtpop3)
        Me.GroupBox1.Controls.Add(Me.txtPassword)
        Me.GroupBox1.Controls.Add(Me.txtUserName)
        Me.GroupBox1.Controls.Add(Me.txtEmailAddress)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(5, 64)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(780, 403)
        Me.GroupBox1.TabIndex = 42
        Me.GroupBox1.TabStop = False
        '
        'txtPortE
        '
        Me.txtPortE.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPortE.Location = New System.Drawing.Point(463, 127)
        Me.txtPortE.Name = "txtPortE"
        Me.txtPortE.Size = New System.Drawing.Size(70, 21)
        Me.txtPortE.TabIndex = 72
        Me.txtPortE.Text = "110"
        '
        'txtAttachmentE
        '
        Me.txtAttachmentE.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAttachmentE.Location = New System.Drawing.Point(463, 154)
        Me.txtAttachmentE.Name = "txtAttachmentE"
        Me.txtAttachmentE.Size = New System.Drawing.Size(302, 21)
        Me.txtAttachmentE.TabIndex = 73
        Me.txtAttachmentE.Text = "D:\New Folder\PASI\POP3\GetEmail\GetEmail\bin\Debug\inbox"
        '
        'txtpop3E
        '
        Me.txtpop3E.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtpop3E.Location = New System.Drawing.Point(463, 100)
        Me.txtpop3E.Name = "txtpop3E"
        Me.txtpop3E.Size = New System.Drawing.Size(302, 21)
        Me.txtpop3E.TabIndex = 71
        Me.txtpop3E.Text = "mail.tos.co.id"
        '
        'txtPasswordE
        '
        Me.txtPasswordE.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPasswordE.Location = New System.Drawing.Point(463, 73)
        Me.txtPasswordE.Name = "txtPasswordE"
        Me.txtPasswordE.Size = New System.Drawing.Size(302, 21)
        Me.txtPasswordE.TabIndex = 70
        Me.txtPasswordE.Text = "iswari"
        Me.txtPasswordE.UseSystemPasswordChar = True
        '
        'txtUserNameE
        '
        Me.txtUserNameE.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUserNameE.Location = New System.Drawing.Point(463, 46)
        Me.txtUserNameE.Name = "txtUserNameE"
        Me.txtUserNameE.Size = New System.Drawing.Size(302, 21)
        Me.txtUserNameE.TabIndex = 69
        Me.txtUserNameE.Text = "dian@tos.co.id"
        '
        'txtEmailAddressE
        '
        Me.txtEmailAddressE.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmailAddressE.Location = New System.Drawing.Point(463, 19)
        Me.txtEmailAddressE.Name = "txtEmailAddressE"
        Me.txtEmailAddressE.Size = New System.Drawing.Size(302, 21)
        Me.txtEmailAddressE.TabIndex = 68
        Me.txtEmailAddressE.Text = "dian@tos.co.id"
        '
        'rtbProcess
        '
        Me.rtbProcess.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rtbProcess.Location = New System.Drawing.Point(6, 208)
        Me.rtbProcess.Name = "rtbProcess"
        Me.rtbProcess.ReadOnly = True
        Me.rtbProcess.Size = New System.Drawing.Size(767, 132)
        Me.rtbProcess.TabIndex = 67
        Me.rtbProcess.Text = ""
        '
        'txtLast
        '
        Me.txtLast.BackColor = System.Drawing.SystemColors.Control
        Me.txtLast.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtLast.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLast.Location = New System.Drawing.Point(162, 346)
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
        Me.txtNext.Location = New System.Drawing.Point(162, 373)
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
        Me.Label8.Location = New System.Drawing.Point(9, 375)
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
        Me.Label7.Location = New System.Drawing.Point(9, 348)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(87, 13)
        Me.Label7.TabIndex = 51
        Me.Label7.Text = "Last Process :"
        '
        'txtPort
        '
        Me.txtPort.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPort.Location = New System.Drawing.Point(144, 127)
        Me.txtPort.Name = "txtPort"
        Me.txtPort.Size = New System.Drawing.Size(70, 21)
        Me.txtPort.TabIndex = 4
        Me.txtPort.Text = "110"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(9, 130)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(27, 13)
        Me.Label9.TabIndex = 49
        Me.Label9.Text = "Port"
        '
        'txtSechedule
        '
        Me.txtSechedule.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSechedule.Location = New System.Drawing.Point(144, 181)
        Me.txtSechedule.Name = "txtSechedule"
        Me.txtSechedule.Size = New System.Drawing.Size(70, 21)
        Me.txtSechedule.TabIndex = 6
        Me.txtSechedule.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtAttachment
        '
        Me.txtAttachment.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAttachment.Location = New System.Drawing.Point(144, 154)
        Me.txtAttachment.Name = "txtAttachment"
        Me.txtAttachment.Size = New System.Drawing.Size(302, 21)
        Me.txtAttachment.TabIndex = 5
        Me.txtAttachment.Text = "D:\New Folder\PASI\POP3\GetEmail\GetEmail\bin\Debug\inbox"
        '
        'txtpop3
        '
        Me.txtpop3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtpop3.Location = New System.Drawing.Point(144, 100)
        Me.txtpop3.Name = "txtpop3"
        Me.txtpop3.Size = New System.Drawing.Size(302, 21)
        Me.txtpop3.TabIndex = 3
        Me.txtpop3.Text = "mail.tos.co.id"
        '
        'txtPassword
        '
        Me.txtPassword.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPassword.Location = New System.Drawing.Point(144, 73)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(302, 21)
        Me.txtPassword.TabIndex = 2
        Me.txtPassword.Text = "iswari"
        Me.txtPassword.UseSystemPasswordChar = True
        '
        'txtUserName
        '
        Me.txtUserName.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUserName.Location = New System.Drawing.Point(144, 46)
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.Size = New System.Drawing.Size(302, 21)
        Me.txtUserName.TabIndex = 1
        Me.txtUserName.Text = "dian@tos.co.id"
        '
        'txtEmailAddress
        '
        Me.txtEmailAddress.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmailAddress.Location = New System.Drawing.Point(144, 19)
        Me.txtEmailAddress.Name = "txtEmailAddress"
        Me.txtEmailAddress.Size = New System.Drawing.Size(302, 21)
        Me.txtEmailAddress.TabIndex = 0
        Me.txtEmailAddress.Text = "dian@tos.co.id"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(9, 184)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(132, 13)
        Me.Label6.TabIndex = 41
        Me.Label6.Text = "Schedule Every (Seconds)"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(9, 157)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(123, 13)
        Me.Label5.TabIndex = 40
        Me.Label5.Text = "Attachment Save Folder"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(9, 103)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(36, 13)
        Me.Label4.TabIndex = 39
        Me.Label4.Text = "POP 3"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(9, 76)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(53, 13)
        Me.Label3.TabIndex = 38
        Me.Label3.Text = "Password"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(9, 49)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 13)
        Me.Label2.TabIndex = 37
        Me.Label2.Text = "User Name"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(9, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 13)
        Me.Label1.TabIndex = 36
        Me.Label1.Text = "E-mail Address"
        '
        'txtcounter
        '
        Me.txtcounter.Location = New System.Drawing.Point(237, 537)
        Me.txtcounter.Name = "txtcounter"
        Me.txtcounter.Size = New System.Drawing.Size(72, 20)
        Me.txtcounter.TabIndex = 48
        Me.txtcounter.Visible = False
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnExit.BackColor = System.Drawing.SystemColors.Control
        Me.btnExit.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.Image = Global.GET_MAIL.My.Resources.Resources.door_out
        Me.btnExit.Location = New System.Drawing.Point(5, 533)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(87, 28)
        Me.btnExit.TabIndex = 10
        Me.btnExit.Text = "   &Exit"
        Me.btnExit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'btnManual
        '
        Me.btnManual.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnManual.BackColor = System.Drawing.SystemColors.Control
        Me.btnManual.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnManual.Image = Global.GET_MAIL.My.Resources.Resources.control_play_blue
        Me.btnManual.Location = New System.Drawing.Point(668, 532)
        Me.btnManual.Name = "btnManual"
        Me.btnManual.Size = New System.Drawing.Size(117, 29)
        Me.btnManual.TabIndex = 8
        Me.btnManual.Text = " &Manual Process"
        Me.btnManual.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnManual.UseVisualStyleBackColor = False
        '
        'timerProcess
        '
        Me.timerProcess.Interval = 2000
        '
        'lblDB
        '
        Me.lblDB.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblDB.Font = New System.Drawing.Font("Verdana", 6.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDB.ForeColor = System.Drawing.Color.Black
        Me.lblDB.Location = New System.Drawing.Point(285, 564)
        Me.lblDB.Name = "lblDB"
        Me.lblDB.Size = New System.Drawing.Size(500, 13)
        Me.lblDB.TabIndex = 54
        Me.lblDB.Text = "Next Process : asdadasdasdsadas"
        Me.lblDB.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'frmGetMail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.BackColor = System.Drawing.Color.LightSteelBlue
        Me.ClientSize = New System.Drawing.Size(790, 579)
        Me.Controls.Add(Me.lblDB)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnManual)
        Me.Controls.Add(Me.grpMsg)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.txtcounter)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmGetMail"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Tag = "GetMail"
        Me.Text = "GET MAIL"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.grpMsg.ResumeLayout(False)
        Me.grpMsg.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Public WithEvents txtMsg As System.Windows.Forms.TextBox
    Public WithEvents grpMsg As System.Windows.Forms.GroupBox
    Friend WithEvents btnManual As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtLast As System.Windows.Forms.TextBox
    Friend WithEvents txtNext As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtPort As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtcounter As System.Windows.Forms.TextBox
    Friend WithEvents txtSechedule As System.Windows.Forms.TextBox
    Friend WithEvents txtAttachment As System.Windows.Forms.TextBox
    Friend WithEvents txtpop3 As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtUserName As System.Windows.Forms.TextBox
    Friend WithEvents txtEmailAddress As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents timerProcess As System.Windows.Forms.Timer
    Friend WithEvents rtbProcess As System.Windows.Forms.RichTextBox
    Friend WithEvents txtPortE As System.Windows.Forms.TextBox
    Friend WithEvents txtAttachmentE As System.Windows.Forms.TextBox
    Friend WithEvents txtpop3E As System.Windows.Forms.TextBox
    Friend WithEvents txtPasswordE As System.Windows.Forms.TextBox
    Friend WithEvents txtUserNameE As System.Windows.Forms.TextBox
    Friend WithEvents txtEmailAddressE As System.Windows.Forms.TextBox
    Friend WithEvents lblDB As System.Windows.Forms.Label

End Class
