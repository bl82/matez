Option Strict Off
Option Explicit On 


Friend Class Form1
    Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
    Public Sub New()
        MyBase.New()
        If m_vb6FormDefInstance Is Nothing Then
            If m_InitializingDefInstance Then
                m_vb6FormDefInstance = Me
            Else
                Try
                    'For the start-up form, the first instance created is the default instance.
                    If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
                        m_vb6FormDefInstance = Me
                    End If
                Catch
                End Try
            End If
        End If
        'This call is required by the Windows Form Designer.
        InitializeComponent()
    End Sub
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public WithEvents Command2 As System.Windows.Forms.Button
    Public WithEvents Text1 As System.Windows.Forms.TextBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Public WithEvents Button1 As System.Windows.Forms.Button
    Public WithEvents Undo As System.Windows.Forms.Button
    Public WithEvents text2 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem13 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem14 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem15 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem16 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem17 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem18 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem19 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem20 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem21 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem22 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem23 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem24 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem26 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem27 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem28 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem29 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem25 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem30 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem31 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem32 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem33 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem34 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem35 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem36 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem37 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem38 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem39 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem40 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem43 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem41 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem42 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem44 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem45 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem46 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem47 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem48 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem49 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem50 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem51 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem52 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem53 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem55 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem56 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem57 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem58 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem59 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem60 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem61 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem62 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem63 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem64 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem65 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem66 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem67 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem68 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem69 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem70 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem71 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem72 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem73 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem74 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem75 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem76 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem77 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem78 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem79 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem80 As System.Windows.Forms.MenuItem
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents MenuItem81 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem82 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem83 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem84 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem85 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem86 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem87 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem54 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem88 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem89 As System.Windows.Forms.MenuItem
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox6 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox7 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox8 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox9 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox10 As System.Windows.Forms.PictureBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents MenuItem90 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem91 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem92 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem93 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem94 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem95 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem96 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem97 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem98 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem99 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem100 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem101 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem102 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem103 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem104 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem105 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem106 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem107 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem108 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem109 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem110 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem111 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem112 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem113 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem114 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem115 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem116 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem117 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem118 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem119 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem120 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem121 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem122 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem123 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem124 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem125 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem126 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem127 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem128 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem129 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem130 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem131 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem132 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem133 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem134 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem135 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem136 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem137 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem138 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem140 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem139 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem141 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem142 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem143 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem144 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem145 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem146 As System.Windows.Forms.MenuItem
    Friend WithEvents sx As System.Windows.Forms.Label
    Friend WithEvents MenuItem147 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem148 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem149 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem150 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.Text1 = New System.Windows.Forms.TextBox
        Me.text2 = New System.Windows.Forms.TextBox
        Me.Command2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.Undo = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem17 = New System.Windows.Forms.MenuItem
        Me.MenuItem18 = New System.Windows.Forms.MenuItem
        Me.MenuItem19 = New System.Windows.Forms.MenuItem
        Me.MenuItem20 = New System.Windows.Forms.MenuItem
        Me.MenuItem80 = New System.Windows.Forms.MenuItem
        Me.MenuItem130 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuItem11 = New System.Windows.Forms.MenuItem
        Me.MenuItem14 = New System.Windows.Forms.MenuItem
        Me.MenuItem21 = New System.Windows.Forms.MenuItem
        Me.MenuItem22 = New System.Windows.Forms.MenuItem
        Me.MenuItem23 = New System.Windows.Forms.MenuItem
        Me.MenuItem26 = New System.Windows.Forms.MenuItem
        Me.MenuItem27 = New System.Windows.Forms.MenuItem
        Me.MenuItem28 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuItem41 = New System.Windows.Forms.MenuItem
        Me.MenuItem56 = New System.Windows.Forms.MenuItem
        Me.MenuItem73 = New System.Windows.Forms.MenuItem
        Me.MenuItem29 = New System.Windows.Forms.MenuItem
        Me.MenuItem16 = New System.Windows.Forms.MenuItem
        Me.MenuItem13 = New System.Windows.Forms.MenuItem
        Me.MenuItem12 = New System.Windows.Forms.MenuItem
        Me.MenuItem24 = New System.Windows.Forms.MenuItem
        Me.MenuItem30 = New System.Windows.Forms.MenuItem
        Me.MenuItem15 = New System.Windows.Forms.MenuItem
        Me.MenuItem31 = New System.Windows.Forms.MenuItem
        Me.MenuItem55 = New System.Windows.Forms.MenuItem
        Me.MenuItem25 = New System.Windows.Forms.MenuItem
        Me.MenuItem9 = New System.Windows.Forms.MenuItem
        Me.MenuItem10 = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuItem75 = New System.Windows.Forms.MenuItem
        Me.MenuItem76 = New System.Windows.Forms.MenuItem
        Me.MenuItem77 = New System.Windows.Forms.MenuItem
        Me.MenuItem78 = New System.Windows.Forms.MenuItem
        Me.MenuItem79 = New System.Windows.Forms.MenuItem
        Me.MenuItem81 = New System.Windows.Forms.MenuItem
        Me.MenuItem82 = New System.Windows.Forms.MenuItem
        Me.MenuItem83 = New System.Windows.Forms.MenuItem
        Me.MenuItem84 = New System.Windows.Forms.MenuItem
        Me.MenuItem128 = New System.Windows.Forms.MenuItem
        Me.MenuItem129 = New System.Windows.Forms.MenuItem
        Me.MenuItem85 = New System.Windows.Forms.MenuItem
        Me.MenuItem86 = New System.Windows.Forms.MenuItem
        Me.MenuItem122 = New System.Windows.Forms.MenuItem
        Me.MenuItem123 = New System.Windows.Forms.MenuItem
        Me.MenuItem124 = New System.Windows.Forms.MenuItem
        Me.MenuItem125 = New System.Windows.Forms.MenuItem
        Me.MenuItem126 = New System.Windows.Forms.MenuItem
        Me.MenuItem127 = New System.Windows.Forms.MenuItem
        Me.MenuItem87 = New System.Windows.Forms.MenuItem
        Me.MenuItem54 = New System.Windows.Forms.MenuItem
        Me.MenuItem88 = New System.Windows.Forms.MenuItem
        Me.MenuItem89 = New System.Windows.Forms.MenuItem
        Me.MenuItem132 = New System.Windows.Forms.MenuItem
        Me.MenuItem131 = New System.Windows.Forms.MenuItem
        Me.MenuItem133 = New System.Windows.Forms.MenuItem
        Me.MenuItem134 = New System.Windows.Forms.MenuItem
        Me.MenuItem136 = New System.Windows.Forms.MenuItem
        Me.MenuItem139 = New System.Windows.Forms.MenuItem
        Me.MenuItem141 = New System.Windows.Forms.MenuItem
        Me.MenuItem142 = New System.Windows.Forms.MenuItem
        Me.MenuItem143 = New System.Windows.Forms.MenuItem
        Me.MenuItem144 = New System.Windows.Forms.MenuItem
        Me.MenuItem145 = New System.Windows.Forms.MenuItem
        Me.MenuItem147 = New System.Windows.Forms.MenuItem
        Me.MenuItem32 = New System.Windows.Forms.MenuItem
        Me.MenuItem33 = New System.Windows.Forms.MenuItem
        Me.MenuItem113 = New System.Windows.Forms.MenuItem
        Me.MenuItem111 = New System.Windows.Forms.MenuItem
        Me.MenuItem121 = New System.Windows.Forms.MenuItem
        Me.MenuItem40 = New System.Windows.Forms.MenuItem
        Me.MenuItem120 = New System.Windows.Forms.MenuItem
        Me.MenuItem112 = New System.Windows.Forms.MenuItem
        Me.MenuItem43 = New System.Windows.Forms.MenuItem
        Me.MenuItem34 = New System.Windows.Forms.MenuItem
        Me.MenuItem35 = New System.Windows.Forms.MenuItem
        Me.MenuItem36 = New System.Windows.Forms.MenuItem
        Me.MenuItem37 = New System.Windows.Forms.MenuItem
        Me.MenuItem38 = New System.Windows.Forms.MenuItem
        Me.MenuItem39 = New System.Windows.Forms.MenuItem
        Me.MenuItem90 = New System.Windows.Forms.MenuItem
        Me.MenuItem91 = New System.Windows.Forms.MenuItem
        Me.MenuItem92 = New System.Windows.Forms.MenuItem
        Me.MenuItem93 = New System.Windows.Forms.MenuItem
        Me.MenuItem94 = New System.Windows.Forms.MenuItem
        Me.MenuItem95 = New System.Windows.Forms.MenuItem
        Me.MenuItem96 = New System.Windows.Forms.MenuItem
        Me.MenuItem97 = New System.Windows.Forms.MenuItem
        Me.MenuItem98 = New System.Windows.Forms.MenuItem
        Me.MenuItem99 = New System.Windows.Forms.MenuItem
        Me.MenuItem100 = New System.Windows.Forms.MenuItem
        Me.MenuItem101 = New System.Windows.Forms.MenuItem
        Me.MenuItem102 = New System.Windows.Forms.MenuItem
        Me.MenuItem103 = New System.Windows.Forms.MenuItem
        Me.MenuItem104 = New System.Windows.Forms.MenuItem
        Me.MenuItem105 = New System.Windows.Forms.MenuItem
        Me.MenuItem106 = New System.Windows.Forms.MenuItem
        Me.MenuItem107 = New System.Windows.Forms.MenuItem
        Me.MenuItem108 = New System.Windows.Forms.MenuItem
        Me.MenuItem109 = New System.Windows.Forms.MenuItem
        Me.MenuItem110 = New System.Windows.Forms.MenuItem
        Me.MenuItem114 = New System.Windows.Forms.MenuItem
        Me.MenuItem115 = New System.Windows.Forms.MenuItem
        Me.MenuItem116 = New System.Windows.Forms.MenuItem
        Me.MenuItem117 = New System.Windows.Forms.MenuItem
        Me.MenuItem118 = New System.Windows.Forms.MenuItem
        Me.MenuItem119 = New System.Windows.Forms.MenuItem
        Me.MenuItem42 = New System.Windows.Forms.MenuItem
        Me.MenuItem44 = New System.Windows.Forms.MenuItem
        Me.MenuItem49 = New System.Windows.Forms.MenuItem
        Me.MenuItem74 = New System.Windows.Forms.MenuItem
        Me.MenuItem137 = New System.Windows.Forms.MenuItem
        Me.MenuItem138 = New System.Windows.Forms.MenuItem
        Me.MenuItem140 = New System.Windows.Forms.MenuItem
        Me.MenuItem45 = New System.Windows.Forms.MenuItem
        Me.MenuItem46 = New System.Windows.Forms.MenuItem
        Me.MenuItem47 = New System.Windows.Forms.MenuItem
        Me.MenuItem48 = New System.Windows.Forms.MenuItem
        Me.MenuItem50 = New System.Windows.Forms.MenuItem
        Me.MenuItem64 = New System.Windows.Forms.MenuItem
        Me.MenuItem69 = New System.Windows.Forms.MenuItem
        Me.MenuItem68 = New System.Windows.Forms.MenuItem
        Me.MenuItem62 = New System.Windows.Forms.MenuItem
        Me.MenuItem60 = New System.Windows.Forms.MenuItem
        Me.MenuItem53 = New System.Windows.Forms.MenuItem
        Me.MenuItem51 = New System.Windows.Forms.MenuItem
        Me.MenuItem65 = New System.Windows.Forms.MenuItem
        Me.MenuItem63 = New System.Windows.Forms.MenuItem
        Me.MenuItem71 = New System.Windows.Forms.MenuItem
        Me.MenuItem61 = New System.Windows.Forms.MenuItem
        Me.MenuItem52 = New System.Windows.Forms.MenuItem
        Me.MenuItem67 = New System.Windows.Forms.MenuItem
        Me.MenuItem66 = New System.Windows.Forms.MenuItem
        Me.MenuItem70 = New System.Windows.Forms.MenuItem
        Me.MenuItem72 = New System.Windows.Forms.MenuItem
        Me.MenuItem135 = New System.Windows.Forms.MenuItem
        Me.MenuItem146 = New System.Windows.Forms.MenuItem
        Me.MenuItem57 = New System.Windows.Forms.MenuItem
        Me.MenuItem58 = New System.Windows.Forms.MenuItem
        Me.MenuItem59 = New System.Windows.Forms.MenuItem
        Me.MenuItem148 = New System.Windows.Forms.MenuItem
        Me.MenuItem149 = New System.Windows.Forms.MenuItem
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label12 = New System.Windows.Forms.Label
        Me.PictureBox8 = New System.Windows.Forms.PictureBox
        Me.PictureBox9 = New System.Windows.Forms.PictureBox
        Me.PictureBox10 = New System.Windows.Forms.PictureBox
        Me.PictureBox7 = New System.Windows.Forms.PictureBox
        Me.PictureBox6 = New System.Windows.Forms.PictureBox
        Me.PictureBox5 = New System.Windows.Forms.PictureBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.PictureBox4 = New System.Windows.Forms.PictureBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.sx = New System.Windows.Forms.Label
        Me.MenuItem150 = New System.Windows.Forms.MenuItem
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Text1
        '
        Me.Text1.AcceptsReturn = True
        Me.Text1.AutoSize = False
        Me.Text1.BackColor = System.Drawing.Color.Silver
        Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text1.Font = New System.Drawing.Font("Lucida Math", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text1.Location = New System.Drawing.Point(6, 40)
        Me.Text1.MaxLength = 0
        Me.Text1.Multiline = True
        Me.Text1.Name = "Text1"
        Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text1.Size = New System.Drawing.Size(346, 155)
        Me.Text1.TabIndex = 0
        Me.Text1.Text = ""
        '
        'text2
        '
        Me.text2.AcceptsReturn = True
        Me.text2.AutoSize = False
        Me.text2.BackColor = System.Drawing.Color.Silver
        Me.text2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.text2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.text2.Font = New System.Drawing.Font("Lucida Math", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.text2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.text2.Location = New System.Drawing.Point(6, 200)
        Me.text2.MaxLength = 0
        Me.text2.Multiline = True
        Me.text2.Name = "text2"
        Me.text2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.text2.Size = New System.Drawing.Size(346, 160)
        Me.text2.TabIndex = 6
        Me.text2.Text = ""
        '
        'Command2
        '
        Me.Command2.BackColor = System.Drawing.Color.Red
        Me.Command2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command2.Font = New System.Drawing.Font("Lucida Math", 8.25!)
        Me.Command2.ForeColor = System.Drawing.Color.White
        Me.Command2.Location = New System.Drawing.Point(6, 9)
        Me.Command2.Name = "Command2"
        Me.Command2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command2.Size = New System.Drawing.Size(66, 19)
        Me.Command2.TabIndex = 2
        Me.Command2.Text = "Count %"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.White
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Button1.Font = New System.Drawing.Font("Lucida Math", 8.25!)
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(288, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Button1.Size = New System.Drawing.Size(60, 19)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Exit @"
        '
        'Undo
        '
        Me.Undo.BackColor = System.Drawing.Color.FromArgb(CType(64, Byte), CType(64, Byte), CType(64, Byte))
        Me.Undo.Cursor = System.Windows.Forms.Cursors.Default
        Me.Undo.Font = New System.Drawing.Font("Lucida Math", 8.25!)
        Me.Undo.ForeColor = System.Drawing.Color.White
        Me.Undo.Location = New System.Drawing.Point(80, 8)
        Me.Undo.Name = "Undo"
        Me.Undo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Undo.Size = New System.Drawing.Size(56, 19)
        Me.Undo.TabIndex = 5
        Me.Undo.Text = "Undo *"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Label2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Label2.Font = New System.Drawing.Font("Lucida Math", 8.25!)
        Me.Label2.ForeColor = System.Drawing.Color.Red
        Me.Label2.Location = New System.Drawing.Point(292, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 19)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Input:"
        Me.Label2.Visible = False
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Label3.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Label3.Font = New System.Drawing.Font("Lucida Math", 8.25!)
        Me.Label3.ForeColor = System.Drawing.Color.Red
        Me.Label3.Location = New System.Drawing.Point(292, 200)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 18)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Output:"
        Me.Label3.Visible = False
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem4, Me.MenuItem26, Me.MenuItem29, Me.MenuItem25, Me.MenuItem32, Me.MenuItem42, Me.MenuItem137, Me.MenuItem45, Me.MenuItem46, Me.MenuItem47, Me.MenuItem48, Me.MenuItem50, Me.MenuItem57, Me.MenuItem148, Me.MenuItem149})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2, Me.MenuItem3, Me.MenuItem17, Me.MenuItem18, Me.MenuItem19, Me.MenuItem20, Me.MenuItem80, Me.MenuItem130})
        Me.MenuItem1.Text = "Insert Symbol"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "Ceiling"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.Text = "Floor"
        '
        'MenuItem17
        '
        Me.MenuItem17.Index = 2
        Me.MenuItem17.Text = "Integer Round"
        '
        'MenuItem18
        '
        Me.MenuItem18.Index = 3
        Me.MenuItem18.Text = "Absolute"
        '
        'MenuItem19
        '
        Me.MenuItem19.Index = 4
        Me.MenuItem19.Text = "Factorial"
        '
        'MenuItem20
        '
        Me.MenuItem20.Index = 5
        Me.MenuItem20.Text = "Double Factorial"
        '
        'MenuItem80
        '
        Me.MenuItem80.Index = 6
        Me.MenuItem80.Text = "Root Operation"
        '
        'MenuItem130
        '
        Me.MenuItem130.Index = 7
        Me.MenuItem130.Text = "Percent"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 1
        Me.MenuItem4.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem5, Me.MenuItem6, Me.MenuItem11, Me.MenuItem14, Me.MenuItem21, Me.MenuItem22, Me.MenuItem23})
        Me.MenuItem4.Text = "Insert Triangle"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 0
        Me.MenuItem5.Text = "Catalan's Triangle"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 1
        Me.MenuItem6.Text = "Bernoulli Triangle"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 2
        Me.MenuItem11.Text = "Pascal Triangle"
        '
        'MenuItem14
        '
        Me.MenuItem14.Index = 3
        Me.MenuItem14.Text = "Clark's Triangle"
        '
        'MenuItem21
        '
        Me.MenuItem21.Index = 4
        Me.MenuItem21.Text = "Bell Triangle"
        '
        'MenuItem22
        '
        Me.MenuItem22.Index = 5
        Me.MenuItem22.Text = "Losanitsch's Triangle"
        '
        'MenuItem23
        '
        Me.MenuItem23.Index = 6
        Me.MenuItem23.Text = "Leibniz Harmonic Triangle"
        '
        'MenuItem26
        '
        Me.MenuItem26.Index = 2
        Me.MenuItem26.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem27, Me.MenuItem28, Me.MenuItem8, Me.MenuItem41, Me.MenuItem56, Me.MenuItem73})
        Me.MenuItem26.Text = "Insert Boolean"
        '
        'MenuItem27
        '
        Me.MenuItem27.Index = 0
        Me.MenuItem27.Text = "IsEven"
        '
        'MenuItem28
        '
        Me.MenuItem28.Index = 1
        Me.MenuItem28.Text = "IsOdd"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 2
        Me.MenuItem8.Text = "AreComprime"
        '
        'MenuItem41
        '
        Me.MenuItem41.Index = 3
        Me.MenuItem41.Text = "IsPrime"
        '
        'MenuItem56
        '
        Me.MenuItem56.Index = 4
        Me.MenuItem56.Text = "IsInteger"
        '
        'MenuItem73
        '
        Me.MenuItem73.Index = 5
        Me.MenuItem73.Text = "Parity"
        '
        'MenuItem29
        '
        Me.MenuItem29.Index = 3
        Me.MenuItem29.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem16, Me.MenuItem13, Me.MenuItem12, Me.MenuItem24, Me.MenuItem30, Me.MenuItem15, Me.MenuItem31, Me.MenuItem55})
        Me.MenuItem29.Text = "Insert Number"
        '
        'MenuItem16
        '
        Me.MenuItem16.Index = 0
        Me.MenuItem16.Text = "Bell Number"
        '
        'MenuItem13
        '
        Me.MenuItem13.Index = 1
        Me.MenuItem13.Text = "Central Hex Number"
        '
        'MenuItem12
        '
        Me.MenuItem12.Index = 2
        Me.MenuItem12.Text = "Central Cubic Number"
        '
        'MenuItem24
        '
        Me.MenuItem24.Index = 3
        Me.MenuItem24.Text = "Pronic Number"
        '
        'MenuItem30
        '
        Me.MenuItem30.Index = 4
        Me.MenuItem30.Text = "Demlo Number"
        '
        'MenuItem15
        '
        Me.MenuItem15.Index = 5
        Me.MenuItem15.Text = "Stirling number of the second kind "
        '
        'MenuItem31
        '
        Me.MenuItem31.Index = 6
        Me.MenuItem31.Text = "Repunit"
        '
        'MenuItem55
        '
        Me.MenuItem55.Index = 7
        Me.MenuItem55.Text = "Eulerian Number"
        '
        'MenuItem25
        '
        Me.MenuItem25.Index = 4
        Me.MenuItem25.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem9, Me.MenuItem10, Me.MenuItem7, Me.MenuItem75, Me.MenuItem87, Me.MenuItem132, Me.MenuItem136, Me.MenuItem139, Me.MenuItem147})
        Me.MenuItem25.Text = "Insert Function"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 0
        Me.MenuItem9.Text = "Binomial Coeficient"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 1
        Me.MenuItem10.Text = "Central Binomial Coefficient"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 2
        Me.MenuItem7.Text = "Greatest Common Divisor"
        '
        'MenuItem75
        '
        Me.MenuItem75.Index = 3
        Me.MenuItem75.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem76, Me.MenuItem77, Me.MenuItem78, Me.MenuItem79, Me.MenuItem81, Me.MenuItem82, Me.MenuItem83, Me.MenuItem84, Me.MenuItem128, Me.MenuItem129, Me.MenuItem85, Me.MenuItem86, Me.MenuItem122, Me.MenuItem123, Me.MenuItem124, Me.MenuItem125, Me.MenuItem126, Me.MenuItem127})
        Me.MenuItem75.Text = "Trigonometric Function"
        '
        'MenuItem76
        '
        Me.MenuItem76.Index = 0
        Me.MenuItem76.Text = "Sine"
        '
        'MenuItem77
        '
        Me.MenuItem77.Index = 1
        Me.MenuItem77.Text = "Cosine"
        '
        'MenuItem78
        '
        Me.MenuItem78.Index = 2
        Me.MenuItem78.Text = "Tangent"
        '
        'MenuItem79
        '
        Me.MenuItem79.Index = 3
        Me.MenuItem79.Text = "Cotangent"
        '
        'MenuItem81
        '
        Me.MenuItem81.Index = 4
        Me.MenuItem81.Text = "Secant"
        '
        'MenuItem82
        '
        Me.MenuItem82.Index = 5
        Me.MenuItem82.Text = "Cosecant"
        '
        'MenuItem83
        '
        Me.MenuItem83.Index = 6
        Me.MenuItem83.Text = "Hiperbolic sine"
        '
        'MenuItem84
        '
        Me.MenuItem84.Index = 7
        Me.MenuItem84.Text = "Hiperbolic cosine"
        '
        'MenuItem128
        '
        Me.MenuItem128.Index = 8
        Me.MenuItem128.Text = "Hiperbolic tangent"
        '
        'MenuItem129
        '
        Me.MenuItem129.Index = 9
        Me.MenuItem129.Text = "Hiperbolic cotangent"
        '
        'MenuItem85
        '
        Me.MenuItem85.Index = 10
        Me.MenuItem85.Text = "Hiperbolic secant"
        '
        'MenuItem86
        '
        Me.MenuItem86.Index = 11
        Me.MenuItem86.Text = "Hiperbolic cosecant"
        '
        'MenuItem122
        '
        Me.MenuItem122.Index = 12
        Me.MenuItem122.Text = "Inverse sine"
        '
        'MenuItem123
        '
        Me.MenuItem123.Index = 13
        Me.MenuItem123.Text = "Inverse cosine"
        '
        'MenuItem124
        '
        Me.MenuItem124.Index = 14
        Me.MenuItem124.Text = "Inverse tangent"
        '
        'MenuItem125
        '
        Me.MenuItem125.Index = 15
        Me.MenuItem125.Text = "Inverse cotangent"
        '
        'MenuItem126
        '
        Me.MenuItem126.Index = 16
        Me.MenuItem126.Text = "Inverse secant"
        '
        'MenuItem127
        '
        Me.MenuItem127.Index = 17
        Me.MenuItem127.Text = "Invers cosecant"
        '
        'MenuItem87
        '
        Me.MenuItem87.Index = 4
        Me.MenuItem87.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem54, Me.MenuItem88, Me.MenuItem89})
        Me.MenuItem87.Text = " Comparison Function"
        '
        'MenuItem54
        '
        Me.MenuItem54.Index = 0
        Me.MenuItem54.Text = "Expression compare"
        '
        'MenuItem88
        '
        Me.MenuItem88.Index = 1
        Me.MenuItem88.Text = "Maximum"
        '
        'MenuItem89
        '
        Me.MenuItem89.Index = 2
        Me.MenuItem89.Text = "Minimum"
        '
        'MenuItem132
        '
        Me.MenuItem132.Index = 5
        Me.MenuItem132.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem131, Me.MenuItem133, Me.MenuItem134})
        Me.MenuItem132.Text = "Digital Function"
        '
        'MenuItem131
        '
        Me.MenuItem131.Index = 0
        Me.MenuItem131.Text = "Digital Root"
        '
        'MenuItem133
        '
        Me.MenuItem133.Index = 1
        Me.MenuItem133.Text = "Digital Summ"
        '
        'MenuItem134
        '
        Me.MenuItem134.Index = 2
        Me.MenuItem134.Text = "Digital Block"
        '
        'MenuItem136
        '
        Me.MenuItem136.Index = 6
        Me.MenuItem136.Text = "Round Function"
        '
        'MenuItem139
        '
        Me.MenuItem139.Index = 7
        Me.MenuItem139.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem141, Me.MenuItem142, Me.MenuItem143, Me.MenuItem144, Me.MenuItem145})
        Me.MenuItem139.Text = "Logarithmic Function"
        '
        'MenuItem141
        '
        Me.MenuItem141.Index = 0
        Me.MenuItem141.Text = "Logarithm"
        '
        'MenuItem142
        '
        Me.MenuItem142.Index = 1
        Me.MenuItem142.Text = "Natural Logarithm"
        '
        'MenuItem143
        '
        Me.MenuItem143.Index = 2
        Me.MenuItem143.Text = "2Base Logarithm"
        '
        'MenuItem144
        '
        Me.MenuItem144.Index = 3
        Me.MenuItem144.Text = "10Base Logarithm"
        '
        'MenuItem145
        '
        Me.MenuItem145.Index = 4
        Me.MenuItem145.Text = "Natural Logarithm from 2"
        '
        'MenuItem147
        '
        Me.MenuItem147.Index = 8
        Me.MenuItem147.Text = "Derivation"
        '
        'MenuItem32
        '
        Me.MenuItem32.Index = 5
        Me.MenuItem32.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem33, Me.MenuItem113, Me.MenuItem111, Me.MenuItem121, Me.MenuItem40, Me.MenuItem120, Me.MenuItem112, Me.MenuItem43, Me.MenuItem34, Me.MenuItem90, Me.MenuItem114})
        Me.MenuItem32.Text = "Insert Convertion"
        '
        'MenuItem33
        '
        Me.MenuItem33.Index = 0
        Me.MenuItem33.Text = "Arabic to Roman"
        '
        'MenuItem113
        '
        Me.MenuItem113.Index = 1
        Me.MenuItem113.Text = "Arabic to Miken"
        '
        'MenuItem111
        '
        Me.MenuItem111.Index = 2
        Me.MenuItem111.Text = "Arabic to Mayan"
        '
        'MenuItem121
        '
        Me.MenuItem121.Index = 3
        Me.MenuItem121.Text = "Arabic to Cyrillic"
        '
        'MenuItem40
        '
        Me.MenuItem40.Index = 4
        Me.MenuItem40.Text = "Roman to Arabic"
        '
        'MenuItem120
        '
        Me.MenuItem120.Index = 5
        Me.MenuItem120.Text = "Miken to Arabic"
        '
        'MenuItem112
        '
        Me.MenuItem112.Index = 6
        Me.MenuItem112.Text = "Number to Text"
        '
        'MenuItem43
        '
        Me.MenuItem43.Index = 7
        Me.MenuItem43.Text = "Change Base"
        '
        'MenuItem34
        '
        Me.MenuItem34.Index = 8
        Me.MenuItem34.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem35, Me.MenuItem36, Me.MenuItem37, Me.MenuItem38, Me.MenuItem39})
        Me.MenuItem34.Text = "Insert Roman Numeral"
        '
        'MenuItem35
        '
        Me.MenuItem35.Index = 0
        Me.MenuItem35.Text = "Roman 10^3"
        '
        'MenuItem36
        '
        Me.MenuItem36.Index = 1
        Me.MenuItem36.Text = "Roman 5*10^3"
        '
        'MenuItem37
        '
        Me.MenuItem37.Index = 2
        Me.MenuItem37.Text = "Roman 10^4"
        '
        'MenuItem38
        '
        Me.MenuItem38.Index = 3
        Me.MenuItem38.Text = "Roman 5*10^4"
        '
        'MenuItem39
        '
        Me.MenuItem39.Index = 4
        Me.MenuItem39.Text = "Roman 10^5"
        '
        'MenuItem90
        '
        Me.MenuItem90.Index = 9
        Me.MenuItem90.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem91, Me.MenuItem92, Me.MenuItem93, Me.MenuItem94, Me.MenuItem95, Me.MenuItem96, Me.MenuItem97, Me.MenuItem98, Me.MenuItem99, Me.MenuItem100, Me.MenuItem101, Me.MenuItem102, Me.MenuItem103, Me.MenuItem104, Me.MenuItem105, Me.MenuItem106, Me.MenuItem107, Me.MenuItem108, Me.MenuItem109, Me.MenuItem110})
        Me.MenuItem90.Text = "Insert Mayan Numeral"
        '
        'MenuItem91
        '
        Me.MenuItem91.Index = 0
        Me.MenuItem91.Text = "Mayan 1"
        '
        'MenuItem92
        '
        Me.MenuItem92.Index = 1
        Me.MenuItem92.Text = "Mayan 2"
        '
        'MenuItem93
        '
        Me.MenuItem93.Index = 2
        Me.MenuItem93.Text = "Mayan 3"
        '
        'MenuItem94
        '
        Me.MenuItem94.Index = 3
        Me.MenuItem94.Text = "Mayan 4"
        '
        'MenuItem95
        '
        Me.MenuItem95.Index = 4
        Me.MenuItem95.Text = "Mayan 5"
        '
        'MenuItem96
        '
        Me.MenuItem96.Index = 5
        Me.MenuItem96.Text = "Mayan 6"
        '
        'MenuItem97
        '
        Me.MenuItem97.Index = 6
        Me.MenuItem97.Text = "Mayan 7"
        '
        'MenuItem98
        '
        Me.MenuItem98.Index = 7
        Me.MenuItem98.Text = "Mayan 8"
        '
        'MenuItem99
        '
        Me.MenuItem99.Index = 8
        Me.MenuItem99.Text = "Mayan 9"
        '
        'MenuItem100
        '
        Me.MenuItem100.Index = 9
        Me.MenuItem100.Text = "Mayan 10"
        '
        'MenuItem101
        '
        Me.MenuItem101.Index = 10
        Me.MenuItem101.Text = "Mayan 11"
        '
        'MenuItem102
        '
        Me.MenuItem102.Index = 11
        Me.MenuItem102.Text = "Mayan 12"
        '
        'MenuItem103
        '
        Me.MenuItem103.Index = 12
        Me.MenuItem103.Text = "Mayan 13"
        '
        'MenuItem104
        '
        Me.MenuItem104.Index = 13
        Me.MenuItem104.Text = "Mayan 14"
        '
        'MenuItem105
        '
        Me.MenuItem105.Index = 14
        Me.MenuItem105.Text = "Mayan 15"
        '
        'MenuItem106
        '
        Me.MenuItem106.Index = 15
        Me.MenuItem106.Text = "Mayan 16"
        '
        'MenuItem107
        '
        Me.MenuItem107.Index = 16
        Me.MenuItem107.Text = "Mayan 17"
        '
        'MenuItem108
        '
        Me.MenuItem108.Index = 17
        Me.MenuItem108.Text = "Mayan 18"
        '
        'MenuItem109
        '
        Me.MenuItem109.Index = 18
        Me.MenuItem109.Text = "Mayan 19"
        '
        'MenuItem110
        '
        Me.MenuItem110.Index = 19
        Me.MenuItem110.Text = "Mayan Zero"
        '
        'MenuItem114
        '
        Me.MenuItem114.Index = 10
        Me.MenuItem114.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem115, Me.MenuItem116, Me.MenuItem117, Me.MenuItem118, Me.MenuItem119})
        Me.MenuItem114.Text = "Insert Miken Numeral"
        '
        'MenuItem115
        '
        Me.MenuItem115.Index = 0
        Me.MenuItem115.Text = "Miken 1"
        '
        'MenuItem116
        '
        Me.MenuItem116.Index = 1
        Me.MenuItem116.Text = "Miken 10"
        '
        'MenuItem117
        '
        Me.MenuItem117.Index = 2
        Me.MenuItem117.Text = "Miken 10^2"
        '
        'MenuItem118
        '
        Me.MenuItem118.Index = 3
        Me.MenuItem118.Text = "Miken 10^3"
        '
        'MenuItem119
        '
        Me.MenuItem119.Index = 4
        Me.MenuItem119.Text = "Miken 10^4"
        '
        'MenuItem42
        '
        Me.MenuItem42.Index = 6
        Me.MenuItem42.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem44, Me.MenuItem49, Me.MenuItem74})
        Me.MenuItem42.Text = "Insert Sequence"
        '
        'MenuItem44
        '
        Me.MenuItem44.Index = 0
        Me.MenuItem44.Text = "Perrin Sequence"
        '
        'MenuItem49
        '
        Me.MenuItem49.Index = 1
        Me.MenuItem49.Text = "Fibonacci Number"
        '
        'MenuItem74
        '
        Me.MenuItem74.Index = 2
        Me.MenuItem74.Text = "Thue-Morse Sequence"
        '
        'MenuItem137
        '
        Me.MenuItem137.Index = 7
        Me.MenuItem137.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem138, Me.MenuItem140, Me.MenuItem150})
        Me.MenuItem137.Text = "Insert Operation"
        '
        'MenuItem138
        '
        Me.MenuItem138.Index = 0
        Me.MenuItem138.Text = "Customized Division"
        '
        'MenuItem140
        '
        Me.MenuItem140.Index = 1
        Me.MenuItem140.Text = "Customized Root"
        '
        'MenuItem45
        '
        Me.MenuItem45.Index = 8
        Me.MenuItem45.Text = "Copy"
        '
        'MenuItem46
        '
        Me.MenuItem46.Index = 9
        Me.MenuItem46.Text = "Paste"
        '
        'MenuItem47
        '
        Me.MenuItem47.Index = 10
        Me.MenuItem47.Text = "Select All"
        '
        'MenuItem48
        '
        Me.MenuItem48.Index = 11
        Me.MenuItem48.Text = "Undo"
        '
        'MenuItem50
        '
        Me.MenuItem50.Index = 12
        Me.MenuItem50.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem64, Me.MenuItem69, Me.MenuItem68, Me.MenuItem62, Me.MenuItem60, Me.MenuItem53, Me.MenuItem51, Me.MenuItem65, Me.MenuItem63, Me.MenuItem71, Me.MenuItem61, Me.MenuItem52, Me.MenuItem67, Me.MenuItem66, Me.MenuItem70, Me.MenuItem72, Me.MenuItem135, Me.MenuItem146})
        Me.MenuItem50.Text = "Insert Constant"
        '
        'MenuItem64
        '
        Me.MenuItem64.Index = 0
        Me.MenuItem64.Text = "Apéry's Constant"
        '
        'MenuItem69
        '
        Me.MenuItem69.Index = 1
        Me.MenuItem69.Text = "Artin's Constant"
        '
        'MenuItem68
        '
        Me.MenuItem68.Index = 2
        Me.MenuItem68.Text = "Catalan's Constant"
        '
        'MenuItem62
        '
        Me.MenuItem62.Index = 3
        Me.MenuItem62.Text = "Champernowne Constant"
        '
        'MenuItem60
        '
        Me.MenuItem60.Index = 4
        Me.MenuItem60.Text = "Euler-Mascheroni Constant"
        '
        'MenuItem53
        '
        Me.MenuItem53.Index = 5
        Me.MenuItem53.Text = "Exponent"
        '
        'MenuItem51
        '
        Me.MenuItem51.Index = 6
        Me.MenuItem51.Text = "Golden Ratio"
        '
        'MenuItem65
        '
        Me.MenuItem65.Index = 7
        Me.MenuItem65.Text = "Glaisher-Kinkelin Constant"
        '
        'MenuItem63
        '
        Me.MenuItem63.Index = 8
        Me.MenuItem63.Text = "Khinchin's Constant"
        '
        'MenuItem71
        '
        Me.MenuItem71.Index = 9
        Me.MenuItem71.Text = "Khinchin-Lévy Constant"
        '
        'MenuItem61
        '
        Me.MenuItem61.Index = 10
        Me.MenuItem61.Text = "Mertens' constant"
        '
        'MenuItem52
        '
        Me.MenuItem52.Index = 11
        Me.MenuItem52.Text = "PI"
        '
        'MenuItem67
        '
        Me.MenuItem67.Index = 12
        Me.MenuItem67.Text = "Plastic Constant"
        '
        'MenuItem66
        '
        Me.MenuItem66.Index = 13
        Me.MenuItem66.Text = "Pythagoras's Constant"
        '
        'MenuItem70
        '
        Me.MenuItem70.Index = 14
        Me.MenuItem70.Text = "Ramanujan Constant"
        '
        'MenuItem72
        '
        Me.MenuItem72.Index = 15
        Me.MenuItem72.Text = "Liouville Constant"
        '
        'MenuItem135
        '
        Me.MenuItem135.Index = 16
        Me.MenuItem135.Text = "Stolarsky-Harborth Constant"
        '
        'MenuItem146
        '
        Me.MenuItem146.Index = 17
        Me.MenuItem146.Text = "Natural Logarithm of 2"
        '
        'MenuItem57
        '
        Me.MenuItem57.Index = 13
        Me.MenuItem57.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem58, Me.MenuItem59})
        Me.MenuItem57.Text = "Insert Condition"
        '
        'MenuItem58
        '
        Me.MenuItem58.Index = 0
        Me.MenuItem58.Text = "Insert IF...THEN...END block"
        '
        'MenuItem59
        '
        Me.MenuItem59.Index = 1
        Me.MenuItem59.Text = "Insert IF...THEN...ELSE...END block"
        '
        'MenuItem148
        '
        Me.MenuItem148.Index = 14
        Me.MenuItem148.Text = "Set Variables"
        '
        'MenuItem149
        '
        Me.MenuItem149.Index = 15
        Me.MenuItem149.Text = "Use Variables (if set)"
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.PictureBox1.BackColor = System.Drawing.Color.White
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(160, 247)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(32, 32)
        Me.PictureBox1.TabIndex = 9
        Me.PictureBox1.TabStop = False
        Me.PictureBox1.Visible = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.PictureBox2.Location = New System.Drawing.Point(155, 243)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(42, 42)
        Me.PictureBox2.TabIndex = 10
        Me.PictureBox2.TabStop = False
        Me.PictureBox2.Visible = False
        '
        'PictureBox3
        '
        Me.PictureBox3.BackColor = System.Drawing.Color.White
        Me.PictureBox3.Location = New System.Drawing.Point(158, 246)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(36, 36)
        Me.PictureBox3.TabIndex = 11
        Me.PictureBox3.TabStop = False
        Me.PictureBox3.Visible = False
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Silver
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.PictureBox8)
        Me.Panel1.Controls.Add(Me.PictureBox9)
        Me.Panel1.Controls.Add(Me.PictureBox10)
        Me.Panel1.Controls.Add(Me.PictureBox7)
        Me.Panel1.Controls.Add(Me.PictureBox6)
        Me.Panel1.Controls.Add(Me.PictureBox5)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.PictureBox4)
        Me.Panel1.Location = New System.Drawing.Point(360, 368)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(347, 355)
        Me.Panel1.TabIndex = 12
        Me.Panel1.Visible = False
        '
        'Label12
        '
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label12.Font = New System.Drawing.Font("Lucida Math", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label12.ForeColor = System.Drawing.Color.White
        Me.Label12.Location = New System.Drawing.Point(144, 288)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(56, 16)
        Me.Label12.TabIndex = 15
        Me.Label12.Text = "HELP"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PictureBox8
        '
        Me.PictureBox8.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox8.Image = CType(resources.GetObject("PictureBox8.Image"), System.Drawing.Image)
        Me.PictureBox8.Location = New System.Drawing.Point(264, 312)
        Me.PictureBox8.Name = "PictureBox8"
        Me.PictureBox8.Size = New System.Drawing.Size(32, 32)
        Me.PictureBox8.TabIndex = 14
        Me.PictureBox8.TabStop = False
        Me.PictureBox8.Visible = False
        '
        'PictureBox9
        '
        Me.PictureBox9.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox9.Image = CType(resources.GetObject("PictureBox9.Image"), System.Drawing.Image)
        Me.PictureBox9.Location = New System.Drawing.Point(32, 312)
        Me.PictureBox9.Name = "PictureBox9"
        Me.PictureBox9.Size = New System.Drawing.Size(32, 32)
        Me.PictureBox9.TabIndex = 13
        Me.PictureBox9.TabStop = False
        Me.PictureBox9.Visible = False
        '
        'PictureBox10
        '
        Me.PictureBox10.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox10.Image = CType(resources.GetObject("PictureBox10.Image"), System.Drawing.Image)
        Me.PictureBox10.Location = New System.Drawing.Point(296, 312)
        Me.PictureBox10.Name = "PictureBox10"
        Me.PictureBox10.Size = New System.Drawing.Size(32, 32)
        Me.PictureBox10.TabIndex = 12
        Me.PictureBox10.TabStop = False
        Me.PictureBox10.Visible = False
        '
        'PictureBox7
        '
        Me.PictureBox7.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox7.Image = CType(resources.GetObject("PictureBox7.Image"), System.Drawing.Image)
        Me.PictureBox7.Location = New System.Drawing.Point(248, 152)
        Me.PictureBox7.Name = "PictureBox7"
        Me.PictureBox7.Size = New System.Drawing.Size(32, 32)
        Me.PictureBox7.TabIndex = 11
        Me.PictureBox7.TabStop = False
        '
        'PictureBox6
        '
        Me.PictureBox6.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox6.Image = CType(resources.GetObject("PictureBox6.Image"), System.Drawing.Image)
        Me.PictureBox6.Location = New System.Drawing.Point(16, 152)
        Me.PictureBox6.Name = "PictureBox6"
        Me.PictureBox6.Size = New System.Drawing.Size(32, 32)
        Me.PictureBox6.TabIndex = 10
        Me.PictureBox6.TabStop = False
        '
        'PictureBox5
        '
        Me.PictureBox5.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox5.Image = CType(resources.GetObject("PictureBox5.Image"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(280, 152)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(32, 32)
        Me.PictureBox5.TabIndex = 9
        Me.PictureBox5.TabStop = False
        '
        'Label11
        '
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label11.Font = New System.Drawing.Font("Lucida Math", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label11.ForeColor = System.Drawing.Color.White
        Me.Label11.Location = New System.Drawing.Point(144, 320)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(64, 16)
        Me.Label11.TabIndex = 8
        Me.Label11.Text = "START"
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Lucida Math", 9.75!)
        Me.Label10.Location = New System.Drawing.Point(48, 160)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(192, 16)
        Me.Label10.TabIndex = 7
        Me.Label10.Text = "Digit Processor MAT-EZ v1.0"
        '
        'Label9
        '
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label9.Font = New System.Drawing.Font("Lucida Math", 8.25!)
        Me.Label9.ForeColor = System.Drawing.Color.Maroon
        Me.Label9.Location = New System.Drawing.Point(216, 80)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(112, 16)
        Me.Label9.TabIndex = 6
        Me.Label9.Text = "evg_zel@mail.ru"
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Lucida Math", 8.25!)
        Me.Label8.Location = New System.Drawing.Point(128, 80)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(72, 16)
        Me.Label8.TabIndex = 5
        Me.Label8.Text = "Mailto:"
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Lucida Math", 8.25!)
        Me.Label7.Location = New System.Drawing.Point(128, 56)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(72, 16)
        Me.Label7.TabIndex = 4
        Me.Label7.Text = "HomePage:"
        '
        'Label6
        '
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label6.Font = New System.Drawing.Font("Lucida Math", 8.25!)
        Me.Label6.ForeColor = System.Drawing.Color.Maroon
        Me.Label6.Location = New System.Drawing.Point(208, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(128, 16)
        Me.Label6.TabIndex = 3
        Me.Label6.Text = "http://ezer0.net/"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Lucida Math", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label5.Location = New System.Drawing.Point(8, 192)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(328, 72)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "Multifunctional Calculator, which interpretes numbers as strings, thus enabling c" & _
        "omputation of large digits. All operations are entered through text field. Inclu" & _
        "des different basic and advanced mathematical functions, costants, operations pl" & _
        "us conditional computations."
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Lucida Math", 11.25!, System.Drawing.FontStyle.Bold)
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(144, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(176, 16)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Division by EZer0"
        '
        'PictureBox4
        '
        Me.PictureBox4.Image = CType(resources.GetObject("PictureBox4.Image"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(16, 16)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(112, 120)
        Me.PictureBox4.TabIndex = 0
        Me.PictureBox4.TabStop = False
        '
        'Label1
        '
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label1.Font = New System.Drawing.Font("Lucida Math", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle))
        Me.Label1.Location = New System.Drawing.Point(216, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 19)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "by EZer0"
        '
        'sx
        '
        Me.sx.AutoSize = True
        Me.sx.Location = New System.Drawing.Point(144, 8)
        Me.sx.Name = "sx"
        Me.sx.Size = New System.Drawing.Size(29, 16)
        Me.sx.TabIndex = 15
        Me.sx.Text = "0 ms"
        '
        'MenuItem150
        '
        Me.MenuItem150.Index = 2
        Me.MenuItem150.Text = "Customized Power"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Gray
        Me.ClientSize = New System.Drawing.Size(358, 365)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Undo)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Command2)
        Me.Controls.Add(Me.Text1)
        Me.Controls.Add(Me.text2)
        Me.Controls.Add(Me.sx)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(4, 23)
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "                               D-I-G-I-T   P-R-O-C-E-S-S-O-R"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As Form1
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As Form1
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New Form1
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal Value As Form1)
            m_vb6FormDefInstance = Value
        End Set
    End Property
#End Region
    Structure variable
        Dim name As String
        Dim value As String
    End Structure
    Public var_list(90) As variable
    Public gb As Byte
    Public gc As Char
    Private Function csc(ByVal a As String, Optional ByVal t As Char = "d") As String
        Dim i, z As Int16
        Dim s As String
        csc = div("1", sin(a, t), 11)
    End Function
    Private Function sec(ByVal a As String, Optional ByVal t As Char = "d") As String
        Dim i, z As Int16
        Dim s As String
        sec = div("1", cos(a, t), 11)
    End Function
    Private Function csch(ByVal a As String, Optional ByVal t As Char = "d") As String
        Dim i, z As Int16
        Dim s As String
        csch = div("1", sinh(a, t), 11)
    End Function
    Private Function sech(ByVal a As String, Optional ByVal t As Char = "d") As String
        Dim i, z As Int16
        Dim s As String
        sech = div("1", cosh(a, t), 11)
    End Function
    Private Function acos(ByVal a As String) As String
        Dim i, z As Int16
        Dim c, s As String
        s = 0
        If Comp(a, "-1") = "a<b" Or Comp(a, "1") = "a>b" Then
            s = "e"
            GoTo endd
        End If

        If Comp(a, "1") = "a=b" Then
            s = "0"
            GoTo endd
        End If

        If Comp(a, "-1") = "a=b" Then
            s = PIP()
            GoTo endd
        End If
        If Comp(a, "0") = "a=b" Then
            s = div(PIP(), 2, 10)
            GoTo endd
        End If
        c = div(PIP, 2, 10)
        s = add(s, c)
        For i = 0 To 30
            '    s = add(s, div(pow(Mul(2, a), Mul(2, i)), Mul(pow(i, 2), BinomCoef(Mul(2, i), i)), 10))
            c = div(Factorial2(subs(Mul(2, i), 1)), Mul(Factorial2(Mul(2, i)), add(Mul(2, i), 1)), 10)
            s = add(s, Mul(c, pow(a, 1 + i * 2)))
        Next i
        '        s = div(s, 2, 10)
        '       s = truncateNull(pow(s, "0.5"))
        ' z = Len(s) - 9
        'If z <= 0 Then
        ' z = Len(s)
        ' End If
endd:
        acos = s 'Mid(s, 1, 10)
    End Function
    Private Function atan(ByVal a As String) As String
        Dim i, z As Int16
        Dim c, s As String
        s = 0
        c = div(PIP, 2, 10)
        s = add(s, c)
        If Comp(a, "=") = "a=b" Then
            s = "0"
            GoTo endd
        End If
        If Comp(a, "-1") = "a>b" And Comp(a, "1") = "a<b" Then
            s = "0"
            GoTo endd
        End If
        s = acot(div(1, a, 10))
endd:
        atan = s 'Mid(s, 1, 10)
    End Function

    Private Function acot(ByVal a As String) As String
        Dim i, z As Int16
        Dim c, s As String
        s = 0
        c = div(PIP, 2, 10)
        s = add(s, c)
        If Comp(a, "-1") = "a<b" Or Comp(a, "1") = "a>b" Then
            s = "0"
            GoTo endd
        End If

        For i = 0 To 30
            '    s = add(s, div(pow(Mul(2, a), Mul(2, i)), Mul(pow(i, 2), BinomCoef(Mul(2, i), i)), 10))
            c = div("1", 1 + i * 2, 10)
            If iseven(i) = "1" Then
                s = subs(s, Mul(c, pow(a, 1 + i * 2)))
            Else
                s = add(s, Mul(c, pow(a, 1 + i * 2)))
            End If
        Next i
        '        s = div(s, 2, 10)
        '       s = truncateNull(pow(s, "0.5"))
        ' z = Len(s) - 9
        'If z <= 0 Then
        ' z = Len(s)
        ' End If
endd:
        acot = s 'Mid(s, 1, 10)
    End Function

    Private Function asin(ByVal a As String) As String
        Dim i, z As Int16
        Dim c, s As String
        s = 0
        If Comp(a, "-1") = "a<b" Or Comp(a, "1") = "a>b" Then
            s = "0"
            GoTo endd
        End If

        If Comp(a, "1") = "a=b" Then
            s = div(PIP(), 2, 10)
            GoTo endd
        End If

        If Comp(a, "-1") = "a=b" Then
            s = "-" & div(PIP(), 2, 10)
            GoTo endd
        End If
        If Comp(a, "-1") = "a=b" Then
            s = "0"
            GoTo endd
        End If
        For i = 0 To 30
            '    s = add(s, div(pow(Mul(2, a), Mul(2, i)), Mul(pow(i, 2), BinomCoef(Mul(2, i), i)), 10))
            c = div(Factorial2(subs(Mul(2, i), 1)), Mul(Factorial2(Mul(2, i)), add(Mul(2, i), 1)), 10)
            s = add(s, Mul(c, pow(a, 1 + i * 2)))
        Next i
        '        s = div(s, 2, 10)
        '       s = truncateNull(pow(s, "0.5"))
        ' z = Len(s) - 9
        'If z <= 0 Then
        ' z = Len(s)
        ' End If
endd:
        asin = s 'Mid(s, 1, 10)
    End Function
    Private Function sin(ByVal a As String, Optional ByVal t As Char = "d") As String
        Dim i, z As Int16
        Dim f As Byte
        Dim s, p2, p, p13, p23, p30, p45, n2, pi As String

        If t = "d" Then
            i = 0
            Do Until Comp(Abso(a), "90") <> "a>b"
                i = i + 1
                If Comp(a, "0") = "a<b" Then
                    a = subs(Abso(a), "180")
                Else
                    a = Mul("-1", subs(Abso(a), "180"))
                End If
            Loop
            Select Case a
                Case "90"
                    s = "1"
                Case "-90"
                    s = "-1"
                Case "60"
                    s = div(Rut(3, 2, 10), 2, 10)
                Case "-60"
                    s = "-" & div(Rut(3, 2, 10), 2, 10)
                Case "45"
                    s = div(Rut(2, 2, 10), 2, 10)
                Case "-45"
                    s = "-" & div(Rut(2, 2, 10), 2, 10)
                Case "30"
                    s = "0.5"
                Case "-30"
                    s = "-0.5"
                Case "0"
                    s = "0"
                Case Else
                    If Comp(a, "0") = "a<b" Then
                        f = 1
                    Else
                        f = 0
                    End If
                    a = Abso(div(Mul(a, PIP), "180", 10))
                    For i = 1 To 30
                        n2 = subs(Mul(i, "2"), "1")
                        If isodd(i - 1) = "1" Then
                            s = add(s, "-" & div(pow(a, n2), Factorial(n2), 20))
                        Else
                            s = add(s, div(pow(a, n2), Factorial(n2), 20))
                        End If
                    Next i
            End Select
        Else
            If Comp(a, "0") = "a<b" Then
                f = 1
            Else
                f = 0
            End If
            p2 = div(PIP(), "2", 10)
            Do Until Comp(Abso(a), 10) <> "a>b"
                i = i + 1
                If Comp(a, "0") = "a<b" Then
                    a = subs(Abso(a), PIP())
                Else
                    a = Mul("-1", subs(Abso(a), PIP()))
                End If
            Loop
            a = Abso(a)
            s = 0
            For i = 1 To 30
                n2 = subs(Mul(i, "2"), "1")
                If isodd(i - 1) = "1" Then
                    s = add(s, "-" & div(pow(a, n2), Factorial(n2), 20))
                Else
                    s = add(s, div(pow(a, n2), Factorial(n2), 20))
                End If
            Next i
        End If
        If f = 0 Then
            sin = s
        Else
            sin = "-" & s
        End If

    End Function
    Private Function cos(ByVal a As String, Optional ByVal t As Char = "d") As String
        Dim i, z, j As Int16
        Dim s, p2, p, p13, p23, p30, p45, pa, n2, pi As String
        If t = "d" Then
            i = 0
            Do Until Comp(Abso(a), "180") <> "a>b"
                j = j + 1
                a = subs(Abso(a), "180")
            Loop

            Select Case Abso(a)
                Case "180"
                    s = "-1"
                Case "90"
                    s = "0"
                Case "120"
                    s = "-0.5"
                Case "60"
                    s = "0.5"
                Case "0"
                    s = "1"
                Case "45"
                    s = div(Rut(2, 2, 10), 2, 10)
                Case "30"
                    s = div(Rut(3, 2, 10), 2, 10)
                Case "150"
                    s = "-" & div(Rut(3, 2, 10), 2, 10)
                Case "135"
                    s = "-" & div(Rut(2, 2, 10), 2, 10)
                Case "30"
                    s = div(Rut(3, 2, 10), 2, 10)
                Case Else
                    a = div(Mul(a, PIP), "180", 10)
                    For i = 1 To 30
                        n2 = Mul(i, "2")
                        If isodd(i) = "1" Then
                            s = add(s, "-" & div(pow(a, n2), Factorial(n2), 20))
                        Else
                            s = add(s, div(pow(a, n2), Factorial(n2), 20))
                        End If

                    Next i
            End Select
        Else
            a = cut(a, 10)
            Do Until Comp(Abso(a), PIP()) <> "a>b"
                i = i + 1
                a = subs(Abso(a), PIP())
            Loop

            a = cut(a, 10)
            s = 1
            For i = 1 To 30
                n2 = Mul(i, "2")
                pa = pow(a, n2)
                If Comp(pa, "0") = "a=b" Then
                    Exit For
                End If
                If isodd(i) = "1" Then
                    s = add(s, "-" & div(pa, Factorial(n2), 20))
                Else
                    s = add(s, div(pa, Factorial(n2), 20))
                End If

            Next i
        End If
        If j > 0 Then
            cos = Mul(s, pow("-1", j))
        Else
            cos = s
        End If
    End Function

    Private Function cut(ByVal a As String, ByVal c As String) As String
        Dim z As Int64
        z = InStr(a, ".")
        If z = 0 Then
            cut = a
        Else
            cut = Mid(a, 1, z) & Mid(a, z + 1, c)
        End If

    End Function

    Private Function cosh(ByVal a As String, Optional ByVal t As Char = "d") As String
        Dim i As Int16
        Dim s, z As String
        If t = "d" Then
            a = cut(div(Mul(a, PIP), "180", 10), 10)
        End If
        If t = "r" Then
            a = a
        End If
        z = pow(Exponent(), a)
        s = div(add(z, div(1, z, 10)), "2", 10)
        cosh = s
    End Function
    Private Function sinh(ByVal a As String, Optional ByVal t As Char = "d") As String
        Dim i As Int16
        Dim s, z As String
        If t = "d" Then
            a = cut(div(Mul(a, PIP), "180", 10), 10)
        End If
        If t = "r" Then
            a = a
        End If
        z = pow(Exponent(), a)
        s = div(subs(z, div(1, z, 10)), "2", 10)
        sinh = s
    End Function
    Private Function tg(ByVal a As String, Optional ByVal t As Char = "d") As String
        Dim a1, b1 As String
        a1 = sin(a, t)
        b1 = cos(a, t)
        tg = div(a1, b1, 10) 'Replace(Math.Tan(Val(a)), ",", ".")
    End Function
    Private Function tanh(ByVal a As String, Optional ByVal t As Char = "d") As String
        tanh = div(sinh(a, t), cosh(a, t), 10) 'Replace(Math.Tan(Val(a)), ",", ".")
    End Function
    Private Function ctgh(ByVal a As String, Optional ByVal t As Char = "d") As String
        ctgh = div(cosh(a, t), sinh(a, t), 10) 'Replace(Math.Tan(Val(a)), ",", ".")
    End Function
    Private Function ctg(ByVal a As String, Optional ByVal t As Char = "d") As String
        ctg = div(cos(a, t), sin(a, t), 10) 'Replace(Math.Tan(Val(a)), ",", ".")
    End Function
    Private Function shorts(ByVal s As String) As String
        Dim i, z As Int64
        If InStr(s, ".") <> 0 Then
            For i = 1 To Len(s)
                If Microsoft.VisualBasic.Right(s, 1) = "0" Then
                    s = Mid(s, 1, Len(s) - 1)
                Else
                    If Microsoft.VisualBasic.Right(s, 1) = "." Then
                        s = Mid(s, 1, Len(s) - 1)
                    End If
                    Exit For
                End If

            Next i
        End If
        If s = "" Then
            z = 1
        Else
            z = 0
            For i = 1 To Len(s)
                If Microsoft.VisualBasic.Mid(s, i, 1) = "0" Then
                    z = z + 1
                Else
                    If Microsoft.VisualBasic.Mid(s, i, 1) = "." Then
                        s = Mid(s, i - 1)
                    Else
                        s = Mid(s, i)
                    End If
                    z = 0
                    Exit For
                End If
            Next i
        End If
        If z <> 0 Then
            s = "0"
        End If
        shorts = s
    End Function
    Private Function ArabicToCyrill(ByVal a As String) As String
        Dim outst As String
        Dim l As Int16
        outst = ""
        l = Len(a)
        If l <= 3 Then
            outst = outst & trigrn(a, 0)
        Else
            If l <= 6 Then
                outst = outst & trigrn(Mid(a, 1, Len(a) - 3), 1)
                outst = outst & trigrn(Mid(a, Len(a) - 2, 3), 0)
            Else
                outst = "This number is too large for Cyrillic number system."
            End If
        End If
        ArabicToCyrill = outst
    End Function

    Private Function trigrn(ByVal s As String, ByVal o As Byte) As String
        Dim z As String
        Dim a As Int16
        s = stringd(3 - Len(s), "0") & s
        a = Mid(s, 1, 1)
        Select Case a
            Case 1
                z = z & ChrW(&H3A1)
            Case 2
                z = z & ChrW(&H421)
            Case 3
                z = z & ChrW(&H3A4)
            Case 4
                z = z & ChrW(&H3BD)
            Case 5
                z = z & ChrW(&H3A6)
            Case 6
                z = z & ChrW(&H3A7)
            Case 7
                z = z & ChrW(&H3A8)
            Case 8
                z = z & ChrW(&H3C9)
            Case 9
                z = z & ChrW(&H426)
        End Select
        If o = 0 Or a = "0" Then
        Else
            z = ChrW(&H2021) & z
        End If
        a = Mid(s, 2, 1)
        Select Case a
            Case 1
                z = z & ChrW(&H399)
            Case 2
                z = z & ChrW(&H39A)
            Case 3
                z = z & ChrW(&H39B)
            Case 4
                z = z & ChrW(&H39C)
            Case 5
                z = z & ChrW(&H39D)
            Case 6
                z = z & ChrW(&H3BE)
            Case 7
                z = z & ChrW(&H39F)
            Case 8
                z = z & ChrW(&H3A0)
            Case 9
                z = z & ChrW(&H427)
        End Select
        If o = 0 Or a = "0" Then
        Else
            z = ChrW(&H2021) & z
        End If
        a = Mid(s, 3, 1)
        Select Case a
            Case 1
                z = z & ChrW(&H391)
            Case 2
                z = z & ChrW(&H392)
            Case 3
                z = z & ChrW(&H393)
            Case 4
                z = z & ChrW(&H414)
            Case 5
                z = z & ChrW(&H395)
            Case 6
                z = z & ChrW(&H53)
            Case 7
                z = z & ChrW(&H396)
            Case 8
                z = z & ChrW(&H397)
            Case 9
                z = z & ChrW(&H398)
        End Select
        If o = 0 Or a = "0" Then
        Else
            z = ChrW(&H2021) & z
        End If
        trigrn = z
    End Function
    Private Function count2(ByVal txt As String, Optional ByVal d As Char = "d") As Object
        Dim z, o As Byte
        Dim i, iu, b, j As Int64
        Dim c, c2 As Char
        Dim s, n, n1, n2 As String

        If InStr(txt, ChrW(282)) > 0 Then
            txt = Replace(txt, ChrW(282), Exponent())
        End If
        If InStr(txt, ChrW(216)) > 0 Then
            txt = Replace(txt, ChrW(216), goldenratio())
        End If
        If InStr(txt, ChrW(928)) > 0 Then
            txt = Replace(txt, ChrW(928), PIP())
        End If
        If InStr(txt, "pi") > 0 Then
            txt = Replace(txt, "pi", PIP())
        End If
        If InStr(UCase(txt), "P") > 0 Then
            txt = Replace(txt, "p", PIP())
            txt = Replace(txt, "P", PIP())
        End If
        If InStr(txt, ChrW(947)) > 0 Then
            txt = Replace(txt, ChrW(947), EMC())
        End If
        If InStr(txt, ChrW(12302)) > 0 Then
            txt = Replace(txt, ChrW(12302), MEC())
        End If
        If InStr(txt, ChrW(12303)) > 0 Then
            txt = Replace(txt, ChrW(12303), CNC())
        End If
        If InStr(txt, ChrW(12304)) > 0 Then
            txt = Replace(txt, ChrW(12304), KCC())
        End If
        If InStr(txt, ChrW(12305)) > 0 Then
            txt = Replace(txt, ChrW(12305), APC())
        End If
        If InStr(txt, ChrW(12306)) > 0 Then
            txt = Replace(txt, ChrW(12306), GKC())
        End If
        If InStr(txt, ChrW(12307)) > 0 Then
            txt = Replace(txt, ChrW(12307), PYC())
        End If
        If InStr(txt, ChrW(12308)) > 0 Then
            txt = Replace(txt, ChrW(12308), PLC())
        End If
        If InStr(txt, ChrW(12309)) > 0 Then
            txt = Replace(txt, ChrW(12309), CAC())
        End If
        If InStr(txt, ChrW(12310)) > 0 Then
            txt = Replace(txt, ChrW(12310), ARC())
        End If
        If InStr(txt, ChrW(12311)) > 0 Then
            txt = Replace(txt, ChrW(12311), RAC())
        End If
        If InStr(txt, ChrW(12312)) > 0 Then
            txt = Replace(txt, ChrW(12312), KLC())
        End If
        If InStr(txt, ChrW(&H3B2)) > 0 Then
            txt = Replace(txt, ChrW(&H3B2), SHC())
        End If
        'Производим все вычисления факториалов
        Do Until InStr(1, txt, "!") = 0
            txt = txt & "@"
            n1 = ""
            o = 0
            n2 = ""
            s = ""
            j = 0
            n = ""
            c = ""
            b = 0
            For i = 1 To Len(txt)
                c = UCase(Mid(txt, i, 1))
                If (Asc(c) > 47 And Asc(c) < 58) Or (Asc(c) > 64 And Asc(c) < 71) Or c = "-" Or c = "O" Or c = "H" Or c = "D" Or c = "B" Then
                    n = n & c
                Else
                    If n1 <> "" And c <> "!" Then
                        n2 = n
                        If InStr(UCase(n2), "H") <> 0 Then
                            n2 = hextodec(n2)
                        End If
                        If InStr(UCase(n2), "B") <> 0 Then
                            n2 = bintodec(n2)
                        End If
                        If InStr(UCase(n2), "O") <> 0 Then
                            n2 = octtodec(n2)
                        End If
                        If InStr(n2, ChrW(&H25)) <> 0 Then
                            n2 = perctonum(n2)
                        End If
                        If o = 1 Then
                            's = Val(n1) ^ Val(n2)
                            s = Factorial(n2)
                            txt = replace_Renamed(Mid(txt, 1, j - 1) & s & Mid(txt, i, Len(txt) - i), ".", ".")
                            Exit For
                        End If
                        If o = 3 Then
                            's = Val(n1) ^ Val(n2)
                            s = Factorial2(n2)
                            txt = replace_Renamed(Mid(txt, 1, j - 2) & s & Mid(txt, i, Len(txt) - i), ".", ".")
                            Exit For
                        End If
                    Else
                        If c = "!" And n1 = "" Then
                            o = "1"
                            n1 = "!"
                            j = i
                            b = j
                        Else
                            If c = "!" Then
                                o = "3"
                                j = i
                            Else
                                j = 0
                            End If
                        End If
                        n = ""
                    End If
                End If


            Next i
        Loop

        'Производим все операции возведения в степень
        Do Until d = "n" Or (InStr(1, txt, "^") = 0 And InStr(1, txt, ChrW(8730)) = 0)
            txt = txt & "@"
            n1 = ""
            o = 0
            n2 = ""
            iu = 0
            s = ""
            j = 0
            n = ""
            c = ""
            b = 0
            For i = 1 To Len(txt)
                c = UCase(Mid(txt, i, 1))
                If (Asc(c) > 47 And Asc(c) < 58) Or (Asc(c) > 64 And Asc(c) < 71) Or (c = "-" And n = "") Or c = "." Or c = "O" Or c = "%" Or c = "H" Or c = "D" Or c = "B" Then
                    n = n & c
                    If j = 0 Then
                        j = i
                    End If
                Else
                    If n1 <> "" Then
                        n2 = n
                        If InStr(UCase(n2), "H") <> 0 Then
                            n2 = hextodec(n2)
                        End If
                        If InStr(UCase(n2), "B") <> 0 Then
                            n2 = bintodec(n2)
                        End If
                        If InStr(UCase(n2), "O") <> 0 Then
                            n2 = octtodec(n2)
                        End If
                        If InStr(n2, ChrW(&H25)) <> 0 Then
                            n2 = perctonum(n2)
                        End If
                        If o = 1 Then
                            's = Val(n1) ^ Val(n2)
                            s = pow(n1, n2)
                            If iu = 0 Then
                                iu = i
                            End If
                            txt = replace_Renamed(Mid(txt, 1, j - 1) & s & Mid(txt, iu, Len(txt) - iu), ".", ".")
                            Exit For
                        End If
                        If o = 2 Then
                            's = Val(n1) ^ Val(n2)
                            s = Rut(n1, n2, 10)
                            If iu = 0 Then
                                iu = i
                            End If
                            txt = replace_Renamed(Mid(txt, 1, j - 1) & s & Mid(txt, iu, Len(txt) - iu), ".", ".")
                            Exit For
                        End If
                    Else
                        If c = "^" Or c = "@" Or c = ChrW(8730) Then
                            n1 = n
                            If InStr(UCase(n1), "H") <> 0 Then
                                n1 = hextodec(n1)
                            End If
                            If InStr(UCase(n1), "B") <> 0 Then
                                n1 = bintodec(n1)
                            End If
                            If InStr(UCase(n1), "O") <> 0 Then
                                n1 = octtodec(n2)
                            End If
                            If InStr(n1, ChrW(&H25)) <> 0 Then
                                n1 = perctonum(n1)
                            End If
                            b = j
                        Else
                            j = 0
                        End If
                        n = ""
                    End If
                End If
                If c = "^" Then
                    o = "1"
                Else
                End If
                If c = ChrW(8730) Then
                    o = "2"
                Else
                End If
            Next i
        Loop

        'Производим все операции умножения и деления
        Do Until InStr(1, txt, "*") = 0 And (InStr(1, txt, "/") = 0 Or d = "n") And InStr(1, txt, "m") = 0 And InStr(1, txt, "\") = 0
            txt = regulate2(txt)
            'txt = txt & "@"
            o = 0
            n1 = ""
            n2 = ""
            s = ""
            j = 0
            n = ""
            c = ""
            b = 0
            For i = 1 To Len(txt)
                c = UCase(Mid(txt, i, 1))
                If (Asc(c) > 47 And Asc(c) < 58) Or (Asc(c) > 64 And Asc(c) < 71) Or c = "." Or c = "O" Or c = "%" Or c = "H" Or c = "D" Or c = "B" Or c = "I" Then
                    n = n & c
                    If j = 0 Then
                        j = i
                    End If
                Else

                    If n1 <> "" Then

                        If c = "-" And Not ((Asc(c2) > 47 And Asc(c2) < 58) Or (Asc(c) > 64 And Asc(c) < 71) Or c = "%" Or c = "O" Or c = "H" Or c = "D" Or c = "B" Or c = "I") Then
                            z = 3
                        Else
                            n2 = n
                            If InStr(n2, "i") <> 0 Then
                                If InStr(n1, "i") <> 0 Then
                                    gb = 0
                                Else
                                    gb = 1
                                End If
                            End If
                            If InStr(UCase(n2), "H") <> 0 Then
                                n2 = hextodec(n2)
                            End If
                            If InStr(UCase(n2), "B") <> 0 Then
                                n2 = bintodec(n2)
                            End If
                            If InStr(UCase(n2), "O") <> 0 Then
                                n2 = octtodec(n2)
                            End If
                            If InStr(n2, ChrW(&H25)) <> 0 Then
                                n2 = perctonum(n2)
                            End If
                            If gb = 1 Then
                                'n1 =  & n1

                            End If
                            If z = 3 Then
                                n2 = "-" & n2
                            End If
                            If o = 1 Then
                                's = Val(n1) * Val(n2)
                                s = Mul(n1, n2)
                                txt = replace_Renamed(Mid(txt, 1, j - 1) & s & Mid(txt, i, Len(txt) - i), ".", ".")
                                Exit For
                            End If
                            If o = 2 And d = "d" Then
                                's = Val(n1) / Val(n2)
                                s = div(n1, n2, 10)
                                If s = "DIVISION BY ZERO!" Then
                                    txt = s
                                Else
                                    txt = replace_Renamed(Mid(txt, 1, j - 1) & s & Mid(txt, i, Len(txt) - i), ".", ".")
                                End If
                                Exit For
                            End If
                            If o = 3 Then
                                's = Val(n1) / Val(n2)
                                s = remn(n1, n2)
                                If s = "DIVISION BY ZERO!" Then
                                    txt = s
                                Else
                                    txt = replace_Renamed(Mid(txt, 1, j - 1) & s & Mid(txt, i, Len(txt) - i), ".", ".")
                                End If
                                Exit For
                            End If
                            If o = 4 Then
                                's = Val(n1) / Val(n2)
                                s = idiv(n1, n2)
                                If s = "DIVISION BY ZERO!" Then
                                    txt = s
                                Else
                                    txt = replace_Renamed(Mid(txt, 1, j - 1) & s & Mid(txt, i, Len(txt) - i), ".", ".")
                                End If
                                Exit For
                            End If
                        End If
                    Else
                        If c = "\" Or c = "M" Or c = "*" Or (c = "/" And d = "d") Or c = "@" Then
                            n1 = n
                            If InStr(n1, "i") <> 0 Then
                                If InStr(n2, "i") <> 0 Then
                                    gb = 0
                                Else
                                    gb = 1
                                End If
                            End If
                            If InStr(UCase(n1), "H") <> 0 Then
                                n1 = hextodec(n1)
                            End If
                            If InStr(UCase(n1), "B") <> 0 Then
                                n1 = bintodec(n1)
                            End If
                            If InStr(UCase(n1), "O") <> 0 Then
                                n1 = octtodec(n1)
                            End If
                            If InStr(n1, ChrW(&H25)) <> 0 Then
                                n1 = perctonum(n1)
                            End If
                            b = j
                        Else
                            j = 0
                        End If
                        n = ""

                    End If
                End If
                If c = "*" Then
                    o = "1"
                Else
                End If
                If c = "M" Then
                    o = "3"
                Else
                End If
                If (c = "/" And d = "d") Then
                    o = "2"
                Else
                End If
                If c = "\" Then
                    o = "4"
                Else
                End If
                If c = "-" Then
                    'z = 3
                Else
                End If
                c2 = c
            Next i
        Loop

        'Производим все операции сложения

        Do Until InStr(1, txt, "+") = 0 And (InStr(2, txt, "-") = 0)
            gb = 0
            txt = regulate2(txt)
            n1 = ""
            n2 = ""
            s = 0
            j = 0
            n = ""
            c = ""
            b = 0
            o = 0
            For i = 1 To Len(txt)
                c = UCase(Mid(txt, i, 1))
                If (Asc(c) > 47 And Asc(c) < 58) Or (Asc(c) > 64 And Asc(c) < 71) Or c = "." Or c = "O" Or c = "%" Or c = "H" Or c = "D" Or c = "B" Then
                    n = n & c
                    If j = 0 Then
                        j = i
                    End If
                Else
                    If n1 <> "" Then
                        n2 = n
                        If InStr(UCase(n2), "H") <> 0 Then
                            n2 = hextodec(n2)
                        End If
                        If InStr(UCase(n2), "B") <> 0 Then
                            n2 = bintodec(n2)
                        End If
                        If InStr(UCase(n2), "O") <> 0 Then
                            n2 = octtodec(n2)
                        End If
                        If InStr(n2, ChrW(&H25)) <> 0 Then
                            n2 = perctonum(n2)
                        End If
                        If o <> 0 Then
                            If o = 2 Then
                                n2 = "-" & n2
                            Else
                            End If
                            If gb = 1 Then
                                n1 = "-" & n1
                            End If
                            's = Val(n1) + Val(n2)
                            s = add(n1, n2)
                            txt = replace_Renamed(Mid(txt, 1, j - 1) & s & Mid(txt, i, Len(txt) - i), ".", ".")
                            Exit For
                        End If
                    Else
                        If c = "-" Or c = "+" Or c = "@" Then
                            If c = "-" And i = 1 Then
                                n = n & c
                                j = i
                            Else
                                n1 = n
                                If InStr(UCase(n1), "H") <> 0 Then
                                    n1 = hextodec(n1)
                                End If
                                If InStr(UCase(n1), "B") <> 0 Then
                                    n1 = bintodec(n1)
                                End If
                                If InStr(UCase(n1), "O") <> 0 Then
                                    n1 = octtodec(n1)
                                End If
                                If InStr(n1, ChrW(&H25)) <> 0 Then
                                    n1 = perctonum(n1)
                                End If
                                b = j
                                n = ""
                            End If
                        Else
                            j = 0
                            n = ""
                        End If

                    End If
                End If
                If c = "+" Then
                    o = "1"
                Else
                End If
                If c = "-" Then
                    o = "2"
                Else
                End If
            Next i
        Loop
        count2 = txt
    End Function
    Private Function regulate2(ByRef txt As Object) As Object
        Dim c As Char
        txt = txt & "@"
        c = Mid(txt, 1, 1)
        If (Asc(c) > 47 And Asc(c) < 58) Or c = "-" Then
            gb = 0
        Else
            ' gb = 1
            gc = c
            'txt = Mid(txt, 2)
        End If
        regulate2 = txt
    End Function
    Public sta As New Stack
    Public sorc As String
    Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
        Dim z, bb, aa As Int64
        Dim z2 As Int64
        sta.Push(Text1.Text)
        Dim src
        If Len(Text1.SelectedText) = 0 Then
            src = Text1.Text
        Else
            src = Text1.SelectedText
        End If
        src = Replace(src, " ", "")
        ' src = pwrupdate(src)

        aa = Date.Now.Ticks

        text2.Text = Maincnt(src)

        bb = Date.Now.Ticks
        sx.Text = (bb - aa) * 10 ^ -4 & " ms"

        PictureBox1.Visible = False
        PictureBox2.Visible = False
        PictureBox3.Visible = False
        text2.Visible = True
    End Sub

    Private Function pwrupdate(ByVal src As String) As String
        Dim i, j, z, p, y As Int64
        Dim c As String
        Dim s, n1, n2, s1, s2 As String
        y = 1
        For i = 1 To Len(src) Step 1
            c = Mid(src, i, 2)
            If c = "^(" Then
                z = InStr(CInt(i + 2), src, ")")
                s = count2(Mid(src, i + 2, z - i - 2), "n")
                z = InStr(1, s, "/")
                n1 = Mid(s, 1, z - 1)
                s1 = "^" & n1
                p = 1
                s2 = Mid(s, Len(n1) + 2) & "/"
                For j = 1 To Len(s2)
                    c = Mid(s2, j, 1)
                    If c = "/" Then
                        n2 = Mid(s2, p, j - p)
                        p = j + 1
                        s1 = s1 & ChrW(8730) & n2
                    End If
                Next
                src = Mid(src, y, i - 1) & s1 & Mid(src, i + Len(s) + 3)
                y = i
            End If
        Next
        pwrupdate = src
    End Function
    Private Function Maincnt(ByVal src As String) As String
        Dim b, e, i, z, z1, z2, zz, sk, fu As Int64
        Dim zu, exp As Boolean
        Dim txt, s, n1, n2, n3, dod, txt2 As String
        Dim c As Char
        Dim idd As Short
begin:
        If src = "DIVISION BY ZERO!" Or src = "" Then
            Maincnt = src
        Else
            If InStr(src, "{") > 0 Then
                z = InStr(src, "_")
                s = Mid(src, 2, z - 2)
                If s = "F0001" Then
                    Maincnt = BaseConv(Mid(src, z + 1, Len(src) - z - 1))
                End If
                If s = "F0002" Then
                    Maincnt = compare(Mid(src, z + 1, Len(src) - z - 1))
                End If
                Exit Function
            End If
            If InStr(UCase(src), "IF") > 0 Then
                z = InStr(UCase(src), "THEN")
                exp = compare(Mid(src, InStr(UCase(src), "IF") + 3, z - InStr(UCase(src), "IF") - 4))
                If exp = True Then
                    e = InStr(UCase(src), " E")
                    dod = Mid(src, z + 5, e - z - 5)
                Else
                    e = InStr(UCase(src), "ELSE")
                    If e = 0 Then
                        dod = "0"
                    Else
                        dod = Mid(src, e + 5, InStr(UCase(src), "END") - e - 6)
                    End If
                End If
                Maincnt = Maincnt(Mid(src, 1, InStr(UCase(src), "IF") - 1) & dod & Mid(src, InStr(UCase(src), "END") + 3))
                Exit Function
            Else
                If InStr(src, "<") > 0 Or InStr(src, ">") > 0 Or (InStr(src, "=") > 0 And InStr(LCase(src), "var") = 0) Then
                    Maincnt = compare(Mid(src, 1))
                    Exit Function
                End If
            End If
            z1 = 0
            z2 = 0
            z = 0
            z = InStr(LCase(src), "#coprime(")
            If z > 0 Then
                idd = 1
                GoTo cnt
            End If
            z = InStr(LCase(src), "#gcd(")
            If z > 0 Then
                idd = 2
                GoTo cnt
            End If
            z = InStr(LCase(src), "#berntriangle(")
            If z > 0 Then
                idd = 3
                GoTo cnt
            End If
            z = InStr(LCase(src), "#cattriangle(")
            If z > 0 Then
                idd = 4
                GoTo cnt
            End If
            z = InStr(LCase(src), "#binomcoef(")
            If z > 0 Then
                idd = 5
                GoTo cnt
            End If
            z = InStr(LCase(src), "#centbicoef(")
            If z > 0 Then
                idd = 6
                GoTo cnt
            End If
            z = InStr(LCase(src), "#pastriangle(")
            If z > 0 Then
                idd = 7
                GoTo cnt
            End If
            z = InStr(LCase(src), "#cencubicnum(")
            If z > 0 Then
                idd = 8
                GoTo cnt
            End If
            z = InStr(LCase(src), "#cenhexnum(")
            If z > 0 Then
                idd = 9
                GoTo cnt
            End If
            z = InStr(LCase(src), "#clarktriangle(")
            If z > 0 Then
                idd = 10
                GoTo cnt
            End If
            z = InStr(LCase(src), "#stirlingn2(")
            If z > 0 Then
                idd = 11
                GoTo cnt
            End If
            z = InStr(LCase(src), "#bellnum(")
            If z > 0 Then
                idd = 12
                GoTo cnt
            End If
            z = InStr(LCase(src), "#belltriangle(")
            If z > 0 Then
                idd = 13
                GoTo cnt
            End If
            z = InStr(LCase(src), "#losatriangle(")
            If z > 0 Then
                idd = 14
                GoTo cnt
            End If
            z = InStr(LCase(src), "#leibharmtriangle(")
            If z > 0 Then
                idd = 15
                GoTo cnt
            End If
            z = InStr(LCase(src), "#pronicnum(")
            If z > 0 Then
                idd = 16
                GoTo cnt
            End If
            z = InStr(LCase(src), "#iseven(")
            If z > 0 Then
                idd = 17
                GoTo cnt
            End If
            z = InStr(LCase(src), "#isodd(")
            If z > 0 Then
                idd = 18
                GoTo cnt
            End If
            z = InStr(LCase(src), "#demlonum(")
            If z > 0 Then
                idd = 19
                GoTo cnt
            End If
            z = InStr(LCase(src), "#repunit(")
            If z > 0 Then
                idd = 20
                GoTo cnt
            End If
            z = InStr(LCase(src), "#toroman(")
            If z > 0 Then
                idd = 21
                GoTo cnt
            End If
            z = InStr(LCase(src), "#fromroman(")
            If z > 0 Then
                idd = 22
                GoTo cnt
            End If
            z = InStr(LCase(src), "#prime(")
            If z > 0 Then
                idd = 23
                GoTo cnt
            End If
            z = InStr(LCase(src), "#perrin(")
            If z > 0 Then
                idd = 24
                GoTo cnt
            End If
            z = InStr(LCase(src), "#fibonacci(")
            If z > 0 Then
                idd = 25
                GoTo cnt
            End If
            z = InStr(LCase(src), "#eulern(")
            If z > 0 Then
                idd = 26
                GoTo cnt
            End If
            z = InStr(LCase(src), "#isinteger(")
            If z > 0 Then
                idd = 27
                GoTo cnt
            End If
            z = InStr(LCase(src), "#sin(")
            If z > 0 Then
                idd = 28
                GoTo cnt
            End If
            z = InStr(LCase(src), "#cos(")
            If z > 0 Then
                idd = 29
                GoTo cnt
            End If
            z = InStr(LCase(src), "#tg(")
            If z > 0 Then
                idd = 30
                GoTo cnt
            End If
            z = InStr(LCase(src), "#ctg(")
            If z > 0 Then
                idd = 31
                GoTo cnt
            End If
            z = InStr(LCase(src), "#sine(")
            If z > 0 Then
                idd = 28
                GoTo cnt
            End If
            z = InStr(LCase(src), "#csin(")
            If z > 0 Then
                idd = 29
                GoTo cnt
            End If
            z = InStr(LCase(src), "#tan(")
            If z > 0 Then
                idd = 30
                GoTo cnt
            End If
            z = InStr(LCase(src), "#ctan(")
            If z > 0 Then
                idd = 31
                GoTo cnt
            End If
            z = InStr(LCase(src), "#liouville(")
            If z > 0 Then
                idd = 32
                GoTo cnt
            End If
            z = InStr(LCase(src), "#parity(")
            If z > 0 Then
                idd = 33
                GoTo cnt
            End If
            z = InStr(LCase(src), "#thuemorse(")
            If z > 0 Then
                idd = 34
                GoTo cnt
            End If
            z = InStr(LCase(src), "#thue-morse(")
            If z > 0 Then
                idd = 34
                GoTo cnt
            End If
            z = InStr(LCase(src), "#digiroot(")
            If z > 0 Then
                idd = 35
                GoTo cnt
            End If
            z = InStr(LCase(src), "#csc(")
            If z > 0 Then
                idd = 36
                GoTo cnt
            End If
            z = InStr(LCase(src), "#cosc(")
            If z > 0 Then
                idd = 36
                GoTo cnt
            End If
            z = InStr(LCase(src), "#cosec(")
            If z > 0 Then
                idd = 36
                GoTo cnt
            End If
            z = InStr(LCase(src), "#csec(")
            If z > 0 Then
                idd = 36
                GoTo cnt
            End If
            z = InStr(LCase(src), "#sec(")
            If z > 0 Then
                idd = 37
                GoTo cnt
            End If
            z = InStr(LCase(src), "#secan(")
            If z > 0 Then
                idd = 37
                GoTo cnt
            End If
            z = InStr(LCase(src), "#sinh(")
            If z > 0 Then
                idd = 38
                GoTo cnt
            End If
            z = InStr(LCase(src), "#sineh(")
            If z > 0 Then
                idd = 38
                GoTo cnt
            End If
            z = InStr(LCase(src), "#cosh(")
            If z > 0 Then
                idd = 39
                GoTo cnt
            End If
            z = InStr(LCase(src), "#csinh(")
            If z > 0 Then
                idd = 39
                GoTo cnt
            End If
            z = InStr(LCase(src), "#csch(")
            If z > 0 Then
                idd = 40
                GoTo cnt
            End If
            z = InStr(LCase(src), "#cosch(")
            If z > 0 Then
                idd = 40
                GoTo cnt
            End If
            z = InStr(LCase(src), "#cosech(")
            If z > 0 Then
                idd = 40
                GoTo cnt
            End If
            z = InStr(LCase(src), "#csech(")
            If z > 0 Then
                idd = 40
                GoTo cnt
            End If
            z = InStr(LCase(src), "#sech(")
            If z > 0 Then
                idd = 41
                GoTo cnt
            End If
            z = InStr(LCase(src), "#secanh(")
            If z > 0 Then
                idd = 41
                GoTo cnt
            End If
            z = InStr(LCase(src), "#max(")
            If z > 0 Then
                idd = 42
                GoTo cnt
            End If
            z = InStr(LCase(src), "#min(")
            If z > 0 Then
                idd = 43
                GoTo cnt
            End If
            z = InStr(LCase(src), "#tomayan(")
            If z > 0 Then
                idd = 44
                GoTo cnt
            End If

            z = InStr(LCase(src), "#numtext(")
            If z > 0 Then
                idd = 45
                GoTo cnt
            End If
            z = InStr(LCase(src), "#tomiken(")
            If z > 0 Then
                idd = 46
                GoTo cnt
            End If
            z = InStr(LCase(src), "#frommiken(")
            If z > 0 Then
                idd = 47
                GoTo cnt
            End If
            z = InStr(LCase(src), "#tocyril(")
            If z > 0 Then
                idd = 48
                GoTo cnt
            End If
            z = InStr(LCase(src), "#tanh(")
            If z > 0 Then
                idd = 49
                GoTo cnt
            End If
            z = InStr(LCase(src), "#ctanh(")
            If z > 0 Then
                idd = 50
                GoTo cnt
            End If
            z = InStr(LCase(src), "#tgh(")
            If z > 0 Then
                idd = 49
                GoTo cnt
            End If
            z = InStr(LCase(src), "#ctgh(")
            If z > 0 Then
                idd = 50
                GoTo cnt
            End If
            z = InStr(LCase(src), "#asin(")
            If z > 0 Then
                idd = 51
                GoTo cnt
            End If
            z = InStr(LCase(src), "#arcsin(")
            If z > 0 Then
                idd = 51
                GoTo cnt
            End If
            z = InStr(LCase(src), "#acos(")
            If z > 0 Then
                idd = 52
                GoTo cnt
            End If
            z = InStr(LCase(src), "#arccos(")
            If z > 0 Then
                idd = 52
                GoTo cnt
            End If
            z = InStr(LCase(src), "#atg(")
            If z > 0 Then
                idd = 53
                GoTo cnt
            End If
            z = InStr(LCase(src), "#atan(")
            If z > 0 Then
                idd = 53
                GoTo cnt
            End If
            z = InStr(LCase(src), "#arctg(")
            If z > 0 Then
                idd = 53
                GoTo cnt
            End If
            z = InStr(LCase(src), "#arctan(")
            If z > 0 Then
                idd = 53
                GoTo cnt
            End If
            z = InStr(LCase(src), "#actg(")
            If z > 0 Then
                idd = 54
                GoTo cnt
            End If
            z = InStr(LCase(src), "#arcctg(")
            If z > 0 Then
                idd = 54
                GoTo cnt
            End If
            z = InStr(LCase(src), "#arccotan(")
            If z > 0 Then
                idd = 54
                GoTo cnt
            End If
            z = InStr(LCase(src), "#asec(")
            If z > 0 Then
                idd = 55
                GoTo cnt
            End If
            z = InStr(LCase(src), "#arcsec(")
            If z > 0 Then
                idd = 55
                GoTo cnt
            End If
            z = InStr(LCase(src), "#acsc(")
            If z > 0 Then
                idd = 56
                GoTo cnt
            End If
            z = InStr(LCase(src), "#arccsc(")
            If z > 0 Then
                idd = 56
                GoTo cnt
            End If
            z = InStr(LCase(src), "#digisum(")
            If z > 0 Then
                idd = 57
                GoTo cnt
            End If
            z = InStr(LCase(src), "#digiblock(")
            If z > 0 Then
                idd = 58
                GoTo cnt
            End If
            z = InStr(LCase(src), "#round(")
            If z > 0 Then
                idd = 59
                GoTo cnt
            End If
            z = InStr(LCase(src), "#div(")
            If z > 0 Then
                idd = 60
                GoTo cnt
            End If
            z = InStr(LCase(src), "#pow(")
            If z > 0 Then
                idd = 61
                GoTo cnt
            End If
            z = InStr(LCase(src), "#root(")
            If z > 0 Then
                idd = 62
                GoTo cnt
            End If
            z = InStr(LCase(src), "#log(")
            If z > 0 Then
                idd = 63
                GoTo cnt
            End If
            z = InStr(LCase(src), "#ln2(")
            If z > 0 Then
                idd = 64
                GoTo cnt
            End If
            z = InStr(LCase(src), "#ln2c(")
            If z > 0 Then
                idd = 65
                GoTo cnt
            End If
            z = InStr(LCase(src), "#ln(")
            If z > 0 Then
                idd = 66
                GoTo cnt
            End If
            z = InStr(LCase(src), "#deriv(")
            If z > 0 Then
                idd = 67
                GoTo cnt
            End If
            z = InStr(LCase(src), "#var(")
            If z > 0 Then
                idd = 68
                GoTo cnt
            End If

            If z = 0 Then
                idd = 0
            End If
cnt:
            b = 0
            fu = 0
            sk = 0
            If z > 0 Then
                For i = z To Len(src)
                    c = Mid(src, i, 1)
                    If c = "#" And b <> 0 And fu = 0 Then
                        fu = i
                    End If
                    If c = "(" And b = 0 Then
                        b = i
                    Else
                        If c = "(" And b <> 0 Then
                            sk = sk + 1
                        End If
                    End If
                    If c = ")" And sk = 0 Then
                        e = i
                        Exit For
                    Else
                        If c = ")" And sk = 1 And fu <> 0 Then
                            txt2 = Maincnt(Mid(src, fu, i - fu + 1))
                            src = Mid(src, 1, fu - 1) & txt2 & Mid(src, i + 1)
                            GoTo begin
                        End If
                        If c = ")" And sk <> 0 Then
                            sk = sk - 1
                        End If

                    End If
                Next
                z1 = e + 1
                z2 = z
                z = InStr(Mid(src, b + 1, e - b - 1), ";")
                If idd = 67 Or idd = 68 Then
                    If z = 0 Then
                        n1 = Mid(src, b + 1, e - b - 1)
                    Else
                        zz = InStr(Mid(src, b + z + 1, e - b - z - 1), ";")
                        If zz = 0 Then
                            n1 = Mid(src, b + 1, z - 1)
                            n2 = Mid(src, b + z + 1, e - b - z - 1)
                        Else
                            n1 = Mid(src, b + 1, z - 1)
                            n2 = Mid(src, b + z + 1, zz - 1)
                            n3 = Mid(src, b + z + zz + 1, e - b - z - zz - 1)
                        End If

                    End If
                Else
                    If z = 0 Then
                        n1 = count2(Mid(src, b + 1, e - b - 1))
                    Else
                        zz = InStr(Mid(src, b + z + 1, e - b - z - 1), ";")
                        If zz = 0 Then
                            n1 = count2(Mid(src, b + 1, z - 1))
                            n2 = count2(Mid(src, b + z + 1, e - b - z - 1))
                        Else
                            n1 = count2(Mid(src, b + 1, z - 1))
                            n2 = count2(Mid(src, b + z + 1, zz - 1))
                            n3 = count2(Mid(src, b + z + zz + 1, e - b - z - zz - 1))
                        End If
                    End If
                End If


                If idd = 1 Then
                    Maincnt = coprime(n1, n2)
                End If
                If idd = 67 Then
                    Maincnt = Derivation(n1, n2)
                End If
                If idd = 17 Then
                    Maincnt = iseven(n1)
                End If
                If idd = 18 Then
                    Maincnt = isodd(n1)
                End If
                If idd = 23 Then
                    Maincnt = Prime(n1)
                End If
                If idd = 33 Then
                    Maincnt = Parity(n1)
                End If
                If idd = 2 Then
                    src = Mid(src, 1, z2 - 1) & GCD2(n1, n2) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 42 Then
                    src = Mid(src, 1, z2 - 1) & Max(n1 & ":") & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 43 Then
                    src = Mid(src, 1, z2 - 1) & Min(n1 & ":") & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 26 Then
                    src = Mid(src, 1, z2 - 1) & eulerN(n1, n2) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 59 Then
                    src = Mid(src, 1, z2 - 1) & RoundMath(n1, n2) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 68 Then
                    src = Mid(src, 1, z2 - 1) & SetVar(n1) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 57 Then
                    If z = 0 Then
                        src = Mid(src, 1, z2 - 1) & digisum(n1) & Mid(src, z1)
                    Else
                        src = Mid(src, 1, z2 - 1) & digisum(n1, n2) & Mid(src, z1)
                    End If
                    GoTo begin
                End If
                If idd = 58 Then
                    If zz = 0 Then
                        src = Mid(src, 1, z2 - 1) & digiblock(n1, n2) & Mid(src, z1)
                    Else
                        src = Mid(src, 1, z2 - 1) & digiblock(n1, n2, n3) & Mid(src, z1)
                    End If
                    GoTo begin
                End If
                If idd = 60 Then
                    If zz = 0 Then
                        src = Mid(src, 1, z2 - 1) & div(n1, n2, 10) & Mid(src, z1)
                    Else
                        src = Mid(src, 1, z2 - 1) & div(n1, n2, n3) & Mid(src, z1)
                    End If
                    GoTo begin
                End If
                If idd = 61 Then
                    If zz = 0 Then
                        src = Mid(src, 1, z2 - 1) & pow(n1, n2, 10) & Mid(src, z1)
                    Else
                        src = Mid(src, 1, z2 - 1) & pow(n1, n2, n3) & Mid(src, z1)
                    End If
                    GoTo begin
                End If
                If idd = 62 Then
                    If zz = 0 Then
                        src = Mid(src, 1, z2 - 1) & Rut(n1, n2, 10) & Mid(src, z1)
                    Else
                        src = Mid(src, 1, z2 - 1) & Rut(n1, n2, n3) & Mid(src, z1)
                    End If
                    GoTo begin
                End If
                If idd = 63 Then
                    If z = 0 Then
                        src = Mid(src, 1, z2 - 1) & Logxb(n1, n2, 10) & Mid(src, z1)
                    Else
                        src = Mid(src, 1, z2 - 1) & Logxb(n1, n2, n3) & Mid(src, z1)
                    End If
                    GoTo begin
                End If

                If idd = 64 Then
                    src = Mid(src, 1, z2 - 1) & Logn2(n1) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 65 Then
                    src = Mid(src, 1, z2 - 1) & ln2() & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 66 Then
                    src = Mid(src, 1, z2 - 1) & Logn(n1) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 3 Then
                    Maincnt = BernTriangle(n1)
                End If
                If idd = 4 Then
                    Maincnt = CatTriangle(n1)
                End If
                If idd = 44 Then
                    Maincnt = ArabicToMayan(n1)
                End If
                If idd = 46 Then
                    Maincnt = ArabictoMiken(n1)
                End If
                If idd = 48 Then
                    Maincnt = ArabicToCyrill(n1)
                End If
                If idd = 5 Then
                    src = Mid(src, 1, z2 - 1) & BinomCoef(n1, n2) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 6 Then
                    src = Mid(src, 1, z2 - 1) & CentBiCoef(n1) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 7 Then
                    Maincnt = PasTriangle(n1)
                End If
                If idd = 21 Then
                    Maincnt = roman(n1)
                End If
                If idd = 45 Then
                    Maincnt = NumSpell(n1)
                End If
                If idd = 22 Then
                    src = Mid(src, 1, z2 - 1) & fromRoman(n1) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 47 Then
                    src = Mid(src, 1, z2 - 1) & MikentoArabic(n1) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 8 Then
                    src = Mid(src, 1, z2 - 1) & CenCubicNum(n1) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 9 Then
                    src = Mid(src, 1, z2 - 1) & CenHexNum(n1) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 10 Then
                    src = ClarkTriangle(n1, n2)
                    GoTo begin
                End If
                If idd = 11 Then
                    src = Mid(src, 1, z2 - 1) & StirlingN2(n1, n2) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 24 Then
                    If z = 0 Then
                        src = Mid(src, 1, z2 - 1) & Perrin(n1) & Mid(src, z1)
                        GoTo begin
                    Else
                        Maincnt = Perrin(n1, n2)
                    End If
                End If
                If idd = 25 Then
                    If z = 0 Then
                        src = Mid(src, 1, z2 - 1) & fibonachi(n1) & Mid(src, z1)
                        GoTo begin
                    Else
                        Maincnt = fibonachi(n1, n2)
                    End If
                End If
                If idd = 34 Then
                    If z = 0 Then
                        src = Mid(src, 1, z2 - 1) & ThueMorse(n1) & Mid(src, z1)
                        GoTo begin
                    Else
                        Maincnt = ThueMorse(n1, n2)
                    End If
                End If
                If idd = 28 Then
                    If z = 0 Then
                        src = Mid(src, 1, z2 - 1) & sin(n1) & Mid(src, z1)
                    Else
                        src = Mid(src, 1, z2 - 1) & sin(n1, n2) & Mid(src, z1)
                    End If
                    GoTo begin
                End If
                If idd = 29 Then
                    If z = 0 Then
                        src = Mid(src, 1, z2 - 1) & cos(n1) & Mid(src, z1)
                    Else
                        src = Mid(src, 1, z2 - 1) & cos(n1, n2) & Mid(src, z1)
                    End If
                    GoTo begin
                End If
                If idd = 30 Then
                    If z = 0 Then
                        src = Mid(src, 1, z2 - 1) & tg(n1) & Mid(src, z1)
                    Else
                        src = Mid(src, 1, z2 - 1) & tg(n1, n2) & Mid(src, z1)
                    End If
                    GoTo begin
                End If
                If idd = 31 Then
                    If z = 0 Then
                        src = Mid(src, 1, z2 - 1) & ctg(n1) & Mid(src, z1)
                    Else
                        src = Mid(src, 1, z2 - 1) & ctg(n1, n2) & Mid(src, z1)
                    End If
                    GoTo begin
                End If
                If idd = 36 Then
                    If z = 0 Then
                        src = Mid(src, 1, z2 - 1) & csc(n1) & Mid(src, z1)
                    Else
                        src = Mid(src, 1, z2 - 1) & csc(n1, n2) & Mid(src, z1)
                    End If
                    GoTo begin
                End If
                If idd = 37 Then
                    If z = 0 Then
                        src = Mid(src, 1, z2 - 1) & sec(n1) & Mid(src, z1)
                    Else
                        src = Mid(src, 1, z2 - 1) & sec(n1, n2) & Mid(src, z1)
                    End If
                    GoTo begin
                End If
                If idd = 38 Then
                    If z = 0 Then
                        src = Mid(src, 1, z2 - 1) & sinh(n1) & Mid(src, z1)
                    Else
                        src = Mid(src, 1, z2 - 1) & sinh(n1, n2) & Mid(src, z1)
                    End If
                    GoTo begin
                End If
                If idd = 39 Then
                    If z = 0 Then
                        src = Mid(src, 1, z2 - 1) & cosh(n1) & Mid(src, z1)
                    Else
                        src = Mid(src, 1, z2 - 1) & cosh(n1, n2) & Mid(src, z1)
                    End If
                    GoTo begin
                End If
                If idd = 40 Then
                    If z = 0 Then
                        src = Mid(src, 1, z2 - 1) & csch(n1) & Mid(src, z1)
                    Else
                        src = Mid(src, 1, z2 - 1) & csch(n1, n2) & Mid(src, z1)
                    End If
                    GoTo begin
                End If

                If idd = 49 Then
                    If z = 0 Then
                        src = Mid(src, 1, z2 - 1) & tanh(n1) & Mid(src, z1)
                    Else
                        src = Mid(src, 1, z2 - 1) & tanh(n1, n2) & Mid(src, z1)
                    End If
                    GoTo begin
                End If
                If idd = 50 Then
                    If z = 0 Then
                        src = Mid(src, 1, z2 - 1) & ctgh(n1) & Mid(src, z1)
                    Else
                        src = Mid(src, 1, z2 - 1) & ctgh(n1, n2) & Mid(src, z1)
                    End If
                    GoTo begin
                End If

                If idd = 51 Then
                    src = Mid(src, 1, z2 - 1) & asin(n1) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 52 Then
                    src = Mid(src, 1, z2 - 1) & acos(n1) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 54 Then
                    src = Mid(src, 1, z2 - 1) & acot(n1) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 53 Then
                    src = Mid(src, 1, z2 - 1) & atan(n1) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 41 Then
                    If z = 0 Then
                        src = Mid(src, 1, z2 - 1) & sech(n1) & Mid(src, z1)
                    Else
                        src = Mid(src, 1, z2 - 1) & sech(n1, n2) & Mid(src, z1)
                    End If
                    GoTo begin
                End If
                If idd = 12 Then
                    src = Mid(src, 1, z2 - 1) & BellNum(n1) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 13 Then
                    Maincnt = BellTriangle(n1)
                End If
                If idd = 14 Then
                    Maincnt = LosaTriangle(n1)
                End If
                If idd = 15 Then
                    Maincnt = LeibHarmTriangle(n1)
                End If
                If idd = 27 Then
                    Maincnt = Mid(src, 1, z2 - 1) & IsInteger(n1) & Mid(src, z1)
                End If
                If idd = 16 Then
                    src = Mid(src, 1, z2 - 1) & PronicNum(n1) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 19 Then
                    src = Mid(src, 1, z2 - 1) & DemloNum(n1) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 20 Then
                    src = Mid(src, 1, z2 - 1) & repunit(n1, n2) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 32 Then
                    src = Mid(src, 1, z2 - 1) & Liouville(n1) & Mid(src, z1)
                    GoTo begin
                End If
                If idd = 35 Then
                    src = Mid(src, 1, z2 - 1) & digiroot(n1) & Mid(src, z1)
                    GoTo begin
                End If
            Else
                If MenuItem149.Checked = True Then
                    src = replaceVar(src)
                End If
                Do Until InStr(1, src, "(") = 0 And InStr(1, src, "|") = 0 And InStr(1, src, "[") = 0 And InStr(1, src, ChrW(12313)) = 0 And InStr(1, src, ChrW(12315)) = 0
                    c = ""
                    b = 0
                    e = 0
                    txt = ""
                    zu = 0
                    For i = 1 To Len(src)
                        c = Mid(src, i, 1)
                        If c = "(" Or c = "[" Or c = ChrW(12313) Or c = ChrW(12315) Then
                            b = i
                        End If
                        If c = "|" And zu = False Then
                            b = i
                            zu = True
                        Else
                            If c = "|" And zu = True Then
                                e = i
                                txt = Mid(src, b + 1, e - b - 1)
                                If b - 1 = 0 And Len(src) < e + 1 Then
                                    src = Abso(count2(txt))
                                Else
                                    src = Mid(src, 1, b - 1) & Abso(count2(txt)) & Mid(src, e + 1)
                                End If
                                zu = False
                                Exit For
                            End If
                        End If
                        If c = ")" Then
                            e = i
                            txt = Mid(src, b + 1, e - b - 1)
                            If b - 1 = 0 And Len(src) < e + 1 Then
                                src = count2(txt)
                            Else
                                src = Mid(src, 1, b - 1) & count2(txt) & Mid(src, e + 1)
                            End If
                            Exit For
                        End If

                        If c = "]" Then
                            e = i
                            txt = Mid(src, b + 1, e - b - 1)
                            If b - 1 = 0 And Len(src) < e + 1 Then
                                src = RCINT(count2(txt))
                            Else
                                src = Mid(src, 1, b - 1) & RCINT(count2(txt)) & Mid(src, e + 1)
                            End If
                            Exit For
                        End If
                        If c = ChrW(12314) Then
                            e = i
                            txt = Mid(src, b + 1, e - b - 1)
                            If b - 1 = 0 And Len(src) < e + 1 Then
                                src = Ceiling(count2(txt))
                            Else
                                src = Mid(src, 1, b - 1) & Ceiling(count2(txt)) & Mid(src, e + 1)
                            End If
                            Exit For
                        End If
                        If c = ChrW(12316) Then
                            e = i
                            txt = Mid(src, b + 1, e - b - 1)
                            If b - 1 = 0 And Len(src) < e + 1 Then
                                src = Floor(count2(txt))
                            Else
                                src = Mid(src, 1, b - 1) & Floor(count2(txt)) & Mid(src, e + 1)
                            End If

                            Exit For
                        End If
                    Next i
                Loop
                Do Until Strings.Right(src, 2) <> Chr(13) & Chr(10)
                    src = Strings.Left(src, Len(src) - 2)
                Loop

                Maincnt = count2(src & "*1")
            End If
        End If
    End Function

    Private Function replaceVar(ByVal s As String) As String
        Dim i As Int16
        Do Until var_list(i).name = ""
            s = Replace(s, var_list(i).name, var_list(i).value)
            i = i + 1
        Loop
        replaceVar = s
    End Function


    Private Function SetVar(ByVal vl As String) As String
        Dim v, s, a As String
        Dim k, b, e, i, z As Int16
        s = vl
        b = 0
        e = 0
        k = 0
        For i = 1 To Len(s)
            a = Mid(s, i, 1)
            If a = "," Then
                e = i
                v = Mid(s, b + 1, e - b - 1)
                z = InStr(v, "=")
                If z <> 0 Then
                    var_list(k).name = Trim(Mid(v, 1, z - 1))
                    var_list(k).value = Trim(Mid(v, z + 1))
                End If
                k = k + 1
                b = i
            End If
        Next i
        e = i
        v = Mid(s, b + 1, e - b - 1)
        z = InStr(v, "=")
        If z <> 0 Then
            var_list(k).name = Trim(Mid(v, 1, z - 1))
            var_list(k).value = Trim(Mid(v, z + 1))
        End If
        k = k + 1
        b = i
        SetVar = ""
    End Function

    Private Function replace_Renamed(ByVal txt1 As String, ByVal txt2 As String, ByVal txt3 As String) As Object
        Dim s As String
        Dim i As Int64
        Dim j As Int64
        j = 1
        For i = 1 To Len(txt1)
            s = Mid(txt1, i, Len(txt2))
            If s = txt2 Then
                txt1 = Mid(txt1, 1, i - 1) & txt3 & Mid(txt1, i + Len(txt2))
                j = i
            End If
        Next i
        replace_Renamed = txt1
    End Function
    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        End
    End Sub

    Private Sub Undo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Undo.Click
        Dim a As Byte
again:
        If sta.Count <> 0 Then
            Text1.Text = sta.Pop
        Else
            a = MsgBox("Stack is empty...", MsgBoxStyle.AbortRetryIgnore, "Digit Processor Error")
        End If
        If a = 4 Then
            GoTo again
        End If
        If a = 3 Then
            End
        End If
    End Sub

    Private Function div(ByVal a As String, ByVal b As String, ByVal n As Integer) As String
        Dim ia, ib As Char
        Dim a1, b1 As Int32
        If Microsoft.VisualBasic.Left(a, 1) = "-" Then
            a = Replace(a, "-", "")
            ia = "-"
        End If
        If Microsoft.VisualBasic.Left(b, 1) = "-" Then
            b = Replace(b, "-", "")
            ib = "-"
        End If
        If InStr(1, a, ".") = 0 Then
            a1 = 0
        Else
            a1 = Len(a) - InStr(1, a, ".")
        End If
        If InStr(1, b, ".") = 0 Then
            b1 = 0
        Else
            b1 = Len(b) - InStr(1, b, ".")
        End If
        a = Replace(a, ".", "")
        b = Replace(b, ".", "")
        If a1 > b1 Then
            b = b & stringd(a1 - b1, "0")
        Else
            a = a & stringd(b1 - a1, "0")
        End If

        If (ia = "-" And ib = "-") Or (ia = Nothing And ib = Nothing) Then
            If Comp(a, b) = "a=b" Then
                div = "1"
            Else
                div = shorts(divi(a, b, n))
            End If
        Else
            If Comp(a, b) = "a=b" Then
                div = "-1"
            Else
                div = "-" & shorts(divi(a, b, n))
            End If
        End If

    End Function
    Private Function repunit(ByVal n As String, ByVal b As String) As String
        repunit = div(subs(pow(b, n), "1"), subs(b, "1"), 10)
    End Function
    Private Function remn(ByVal a As String, ByVal b As String) As String
        Dim s As String
        s = add(a, Mul("-" & b, idiv(a, b)))
        If s = "-0" Then
            s = "0"
        End If
        If subs(div(a, b, 10), idiv(a, b)) = "0" Then
            s = "0"
        End If
        remn = s
    End Function
    Private Function subs(ByVal a As String, ByVal b As String) As String
        If a = b Then
            subs = "0"
        Else
            subs = add(a, "-" & b)
        End If

    End Function

    Private Function idiv(ByVal a As String, ByVal b As String) As String
        idiv = IntPart(div(a, b, 10))
    End Function
    Private Function Liouville(ByVal b As String) As String
        Dim i As Int64
        Dim n As String
        If b = "1" Then
            n = "0.1"
        Else
            n = "0.11"
            For i = 3 To CInt(b)
                n = n & stringd(subs(Factorial(Trim(Str(i))), subs(Trim(Str(Len(n))), "1")), "0") & "1"
            Next i
        End If
        Liouville = n
    End Function
    Private Function ArabictoMiken(ByVal s As String) As String
        Dim i, j, a As Int16
        Dim z As String
        If Len(s) > 6 Or Val(s) > 100000 Then
            ArabictoMiken = "Number is too Large for Miken number system"
        Else
            For i = Len(s) To 1 Step -1
                a = Val(Mid(s, Len(s) - i + 1, 1))
                For j = 1 To a
                    Select Case i
                        Case 6
                            z = stringd(10, ChrW(12348))
                        Case 5
                            z = z & ChrW(12348)
                        Case 4
                            z = z & ChrW(12346)
                        Case 3
                            z = z & ChrW(12345)
                        Case 2
                            z = z & ChrW(12344)
                        Case 1
                            z = z & ChrW(12343)
                    End Select

                Next
            Next
            ArabictoMiken = z
        End If

    End Function
    Private Function MikentoArabic(ByVal s As String) As String
        Dim i, j As Int16
        Dim z As String
        Dim a As Char
        For i = 1 To Len(s)
            a = Mid(s, i, 1)
            Select Case a
                Case ChrW(&H303C)
                    z = add(z, "10000")
                Case ChrW(&H303A)
                    z = add(z, "1000")
                Case ChrW(&H3039)
                    z = add(z, "100")
                Case ChrW(&H3038)
                    z = add(z, "10")
                Case ChrW(&H3037)
                    z = add(z, "1")
            End Select
        Next
        MikentoArabic = z
    End Function

    Private Function ArabicToMayan(ByVal z As String) As String
        Dim r, m, st As String
        Dim a(1) As Byte
        Dim y, yb, i As Int32
        m = z
        yb = ipart(Trim(Str(Logxb(z, 20))))
        If m >= 360 And m < 400 Then
            yb = yb + 1
        Else
            yb = yb
        End If
        ReDim a(Val(yb))
        Do Until m <= 0
            If m >= 360 Then
                y = ipart(Trim(Str(Logxb(m, 20))))
                If m < 400 Then
                    y = y + 1
                End If
                r = 360 * pow(20, y - 2)
rept2:
                m = subs(m, r)
                a(Val(y)) = a(Val(y)) + 1
                If Comp(m, r) = "a>b" Then
                    GoTo rept2
                End If
            Else
                y = ipart(Trim(Str(Logxb(m, 20))))
                r = pow(20, y)
rept:
                m = subs(m, r)
                a(Val(y)) = a(Val(y)) + 1
                If Comp(m, r) <> "a<b" Then
                    GoTo rept
                End If
            End If
        Loop

        For i = 0 To Val(yb)
            If a(i) = Nothing Then
                st = Maya_acnv(0) & st
            Else
                st = Maya_acnv(a(i)) & st
            End If
        Next
        'convertN = st
        ArabicToMayan = st
    End Function

    Private Function NumSpell(ByVal a As String) As String
        Dim s, sm, outst As String
        Dim l, l2, i, n, n2, n3 As Int16
        outst = ""
        's = Trim(Str(a))
        s = a
        l = Len(s)
        i = l - (l \ 3) * 3
        If i <> 0 Then
            s = stringd(3 - i, "0") & s
            l = 3 - i + l
        End If
        l2 = l

        Do While l > 0
            sm = Mid(s, l2 - l + 1, 3)
            If l <= 3 Then
                outst = outst & TrianNum(sm, 0)
            Else
                Select Case l
                    Case 6
                        outst = outst & TrianNum(sm, 1)
                        n = Strings.Right(sm, 1)
                        n2 = Mid(sm, 2, 1)
                        n3 = Mid(sm, 1, 1)
                        If n3 = 0 And n2 = 0 And n = 0 Then
                            outst = outst
                        Else
                            If n = "1" And n2 <> "1" Then
                                outst = outst & "тысяча "
                            Else
                                If (n = "2" Or n = "3" Or n = "4") And n2 <> "1" Then
                                    outst = outst & "тысячи "
                                Else
                                    outst = outst & "тысяч "
                                End If
                            End If
                        End If
                    Case 9
                        outst = outst & TrianNum(sm, 2)
                        n = Strings.Right(sm, 1)
                        n2 = Mid(sm, 2, 1)
                        n3 = Mid(sm, 1, 1)
                        If n3 = 0 And n2 = 0 And n = 0 Then
                            outst = outst
                        Else
                            If n = "1" And n2 <> "1" Then
                                outst = outst & "миллион "
                            Else
                                If (n = "2" Or n = "3" Or n = "4") And n2 <> "1" Then
                                    outst = outst & "миллиона "
                                Else
                                    outst = outst & "миллионов "
                                End If
                            End If
                        End If
                    Case 12
                        outst = outst & TrianNum(sm, 2)
                        n = Strings.Right(sm, 1)
                        n2 = Mid(sm, 2, 1)
                        n3 = Mid(sm, 1, 1)
                        If n3 = 0 And n2 = 0 And n = 0 Then
                            outst = outst
                        Else
                            If n = "1" And n2 <> "1" Then
                                outst = outst & "миллиард "
                            Else
                                If (n = "2" Or n = "3" Or n = "4") And n2 <> "1" Then
                                    outst = outst & "миллиарда "
                                Else
                                    outst = outst & "миллиардов "
                                End If
                            End If
                        End If
                    Case Else
                        outst = outst & TrianNum(sm, 3)
                        n = Strings.Right(sm, 1)
                        n2 = Mid(sm, 2, 1)
                        n3 = Mid(sm, 1, 1)
                        If n3 = 0 And n2 = 0 And n = 0 Then
                            outst = outst
                        Else
                            Select Case l
                                Case 15
                                    outst = outst & "три"
                                Case 18
                                    outst = outst & "квадри"
                                Case 21
                                    outst = outst & "квинти"
                                Case 24
                                    outst = outst & "сексти"
                                Case 27
                                    outst = outst & "септи"
                                Case 30
                                    outst = outst & "окти"
                                Case 33
                                    outst = outst & "нони"
                                Case 36
                                    outst = outst & "деци"
                                Case 39
                                    outst = outst & "ундеци"
                                Case 42
                                    outst = outst & "додеци"
                                Case 45
                                    outst = outst & "тредеци"
                                Case 48
                                    outst = outst & "кваттуордеци"
                                Case 51
                                    outst = outst & "квиндеци"
                                Case 54
                                    outst = outst & "седеци"
                                Case 57
                                    outst = outst & "септдеци"
                                Case 60
                                    outst = outst & "дуодевигинти"
                                Case 63
                                    outst = outst & "ундевигинти"
                                Case 66
                                    outst = outst & "вигинти"
                                Case Else
                                    outst = "Sorry. The number is too large."
                            End Select

                            If n = "1" And n2 <> "1" Then
                                outst = outst & "ллион "
                            Else
                                If (n = "2" Or n = "3" Or n = "4") And n2 <> "1" Then
                                    outst = outst & "ллиона "
                                Else
                                    outst = outst & "ллионов "
                                End If
                            End If
                        End If
                End Select
            End If
            l = l - 3
        Loop
        NumSpell = outst
    End Function

    Private Function TrianNum(ByVal s As String, ByVal j As Byte) As String
        Dim out As String
        Dim c As Char
        out = ""
        c = Mid(s, 1, 1)
        Select Case c
            Case "1"
                out = out + "сто "
            Case "2"
                out = out + "двести "
            Case "3"
                out = out + "триста "
            Case "4"
                out = out + "четыреста "
            Case "5"
                out = out + "пятьсот "
            Case "6"
                out = out + "шестьсот "
            Case "7"
                out = out + "семьсот "
            Case "8"
                out = out + "восемьсот "
            Case "9"
                out = out + "девятьсот "
            Case "0"
                out = out
        End Select
        c = Mid(s, 2, 1)
        If c = "1" Then
            c = Mid(s, 3, 1)
            Select Case c
                Case "1"
                    out = out + "одиннадцать "
                Case "2"
                    out = out + "двенадцать "
                Case "3"
                    out = out + "тринадцать "
                Case "4"
                    out = out + "четырнадцать "
                Case "5"
                    out = out + "пятнадцать "
                Case "6"
                    out = out + "шестнадцать "
                Case "7"
                    out = out + "семнадцать "
                Case "8"
                    out = out + "восемнадцать "
                Case "9"
                    out = out + "девятнадцать "
                Case "0"
                    out = out + "десять"
            End Select
        Else
            c = Mid(s, 2, 1)
            Select Case c
                Case "2"
                    out = out + "двадцать "
                Case "3"
                    out = out + "тридцать "
                Case "4"
                    out = out + "сорок "
                Case "5"
                    out = out + "пятьдесят "
                Case "6"
                    out = out + "шестьдесят "
                Case "7"
                    out = out + "семьдесят "
                Case "8"
                    out = out + "восемьдесят "
                Case "9"
                    out = out + "девяносто "
                Case "0"
                    out = out
            End Select
            c = Mid(s, 3, 1)
            Select Case c
                Case "1"
                    If j = 1 Then
                        out = out + "одна "
                    Else
                        out = out + "один "
                    End If
                Case "2"
                    If j = 1 Then
                        out = out + "две "
                    Else
                        out = out + "два "
                    End If
                Case "3"
                    out = out + "три "
                Case "4"
                    out = out + "четыре "
                Case "5"
                    out = out + "пять "
                Case "6"
                    out = out + "шесть "
                Case "7"
                    out = out + "семь "
                Case "8"
                    out = out + "восемь "
                Case "9"
                    out = out + "девять "
                Case "0"
                    out = out
            End Select
        End If
        TrianNum = out
    End Function

    Private Function KLC() As String
        KLC = "3.27582291872181115978768188245384386360847552598237414940519892419072321564496035518127754047917452949269"
    End Function
    Private Function RAC() As String
        RAC = "262537412640768743.999999999999250072597198185688879353856337336990862707537410378210647910118607312951181"
    End Function
    Private Function SHC() As String
        SHC = "0.81255655901600638769488210164953671243441922490636156678320347580366003142762953508246848982797937869"
    End Function
    Private Function ARC() As String
        ARC = "0.37395581361920228805472805434641641511162924860615004209474280241735018204002808234430431708725056"
    End Function
    Private Function CAC() As String
        CAC = "0.915965594177219015054603514932384110774149374281672134266498119621763019776254769479356512926115106248574"
    End Function
    Private Function PLC() As String
        PLC = "1.32471795724474602596090885447809734073440405690173336453401505030282785124554759405469934798178728032991"
    End Function
    Private Function PYC() As String
        PYC = "1.41421356237309504880168872420969807856967187537694807317667973799073247846210703885038753432764157"
    End Function
    Private Function GKC() As String
        GKC = "1.28242712910062263687534256886979172776768892732500119206374002174040630885882646112973649195820237441024"
    End Function
    Private Function APC() As String
        APC = "1.20205690315959428539973816151144999076498629234049888179227155534183820578631309018645587360933525814619915"
    End Function
    Private Function CNC() As String
        CNC = "0.8622401258680545715577902832493945785657647427682990945160712145573067405905164580420384414386181334"
    End Function
    Private Function KCC() As String
        KCC = "2.68545200106530644530971483548179569382038229399446295305115234555721885953715200280114117493184769799515"
    End Function
    Private Function MEC() As String
        MEC = "0.261497212847642783755426838608695859051566648261199206192064213924924510897368209714142631434246651051617"
    End Function
    Private Function EMC() As String
        EMC = "0.577215664901532860606512090082402431042159335939923598805767234884867726777664670936947063291746749"
    End Function
    Private Function Max(ByVal n As String) As String
        Dim i, y As Int64
        Dim c As Char
        Dim r1, rm As String
        rm = "0"
        y = 1
        For i = 1 To Len(n)
            c = Mid(n, i, 1)
            If c = ":" Then
                r1 = Mid(n, y, i - y)
                y = i + 1
                If Comp(rm, r1) = "a<b" Then
                    rm = r1
                End If
            End If
        Next
        Max = rm
    End Function


    Private Function Min(ByVal n As String) As String
        Dim i, y As Int64
        Dim c As Char
        Dim r1, rm As String
        rm = "0"
        y = 1
        For i = 1 To Len(n)
            c = Mid(n, i, 1)
            If c = ":" Then
                r1 = Mid(n, y, i - y)
                y = i + 1
                If Comp(rm, r1) = "a>b" Or rm = "0" Then
                    rm = r1
                End If
            End If
        Next
        Min = rm
    End Function
    Private Function PIP() As String
        PIP = "3.14159265358979323846264338327950288419716939937510582097494459230781640628620899862803482534211706798214"
    End Function
    Private Function ipart(ByVal n As String) As String
        Dim r As Int32
        r = InStr(1, n, ".")
        If r = 0 Then
            ipart = n
        Else
            ipart = Mid(n, 1, r - 1)
        End If
    End Function
    Private Function fpart(ByVal n As String) As String
        Dim r As Int32
        r = InStr(1, n, ".")
        If r = 0 Then
            fpart = "0"
        Else
            fpart = "0." & Mid(n, r + 1)
        End If
    End Function
    Private Function trim0(ByVal n As String) As String
        Dim i, z As Int32
        Dim c As Char
        For i = 1 To Len(n)
            c = Mid(n, i, 1)
            If c <> "0" Then
                Exit For
            Else
                z = z + 1
            End If
        Next
        If Mid(n, z + 1) = "" Then
            trim0 = "0"
        Else
            trim0 = Mid(n, z + 1)
        End If

    End Function

    Private Function divi(ByVal a As String, ByVal b As String, ByVal n As Integer) As String
        Dim s, re, r, ta, la, ra As String
        Dim i, l, we, le, li, li2, lu, z As Int64
        Dim m As Byte

        a = trim0(a)
        b = trim0(b)
        If a = "0" Or a = "0.0" Then
            s = 0
            GoTo entd
        End If
        If b = "0" Or b = "0.0" Then
            s = "DIVISION BY ZERO!"
            GoTo entd
        End If
        ra = a
        ta = a
        li = 0
        li2 = 0
        'Упрошаем дальнейшие операции подготавливая числа
        l = Len(b)
        If Comp(a, b) = "a>b" Then
            le = l
            For i = 1 To l + 1
                la = Mid(ra, 1, le)
                If Comp(la, b) = "a<b" Then
                    le = le + 1
                Else
                    we = Len(a) - le
                    Exit For
                End If
            Next

        Else
            If Comp(a, b) = "a=b" Then
                s = 1
                z = 0
                GoTo entd
            Else
                For i = 1 To l + 1
                    s = s & "0"
                    ra = ra & "0"
                    z = z + 1
                    If Comp(ra, b) = "a>b" Then
                        la = ra
                        we = 0
                        Exit For
                    End If
                Next i
            End If
        End If
        m = 0
        Do Until z >= n + 1
            ' If truncateNull(la) <> la Then
            'li = li + Len(truncateNull(la))
            'End If
            li = li2
            'Вычисляем наибольший делитель
            For i = 1 To 11
                r = Mul(b, Trim(Str(i)))
                If Comp(r, la) = "a>b" Then
                    m = i - 1
                    Exit For
                End If
            Next
            'Деление
            re = subs(la, subs(r, b))
            'le = Len(re) 'Здесь было le = Len(la), считало неправильно на числе 722/2
            s = s & m
            lu = 1
            If we = 0 Then
                If Comp(re, "0") = "a=b" Then
                    GoTo entd
                End If
                la = re
lend:
                For i = 1 To Len(b) + 1
                    la = la & "0"
                    z = z + 1
                    If lu >= 2 Then
                        s = s & "0"
                    End If
                    If Comp(la, b) = "a<b" Then
                        lu = lu + 1
                    Else
                        If Comp(la, b) = "a=b" Then
                            s = s & "1"
                            GoTo entd
                        End If
                        Exit For
                    End If
                Next
            Else
                If truncateNull(Mid(a, Len(la) + 1)) = 0 And re = "0" Then
                    s = s & stringd(we, "0")
                    GoTo entd
                End If
                li = li + 1
                li2 = li
                For i = 1 To Len(Mid(a, Len(la) + 1))
                    'if we="1" and b=
                    If Comp(re, "0") = "a=b" Then
                        ' s = s & stringd(Len(b) - 1, "0")
                        ' ' li = li + 1
                        re = ""
                    End If
                    la = re & Mid(a, le + li, lu)

                    we = we - 1
                    If Comp(la, b) = "a<b" Then
                        'If we = 0 Then GoTo entd
                        lu = lu + 1
                        If lu >= 2 Then
                            s = s & "0"
                            li2 = li2 + 1
                        End If
                    Else
                        Exit For
                    End If
                    If we = 0 Then
                        If la = "0" Then
                            s = s & la
                            GoTo entd
                        Else
                            lu = 1
                            GoTo lend
                        End If
                    End If
                Next i

            End If
        Loop
entd:
        If z > n Then
            z = z - 1
        End If
        divi = Microsoft.VisualBasic.Left(s, Len(s) - z) & "." & Microsoft.VisualBasic.Right(s, z)
    End Function

    Private Function Perrin(ByVal k As String, Optional ByVal c As Char = "0") As String
        Dim f(1), i As Int64
        Dim s As String
        ReDim f(CInt(k))
        f(0) = 3
        f(1) = 0
        f(2) = 2
        For i = 3 To CInt(k)
            f(i) = f(i - 2) + f(i - 3)
        Next i
        If c = "1" Then
            For i = 0 To CInt(k)
                s = s & Trim(Str(f(i))) & "."
            Next
            Perrin = s
        Else
            Perrin = Trim(Str(f(CInt(k))))
        End If

    End Function
    Private Function ThueMorse(ByVal k As String, Optional ByVal c As Char = "0") As String
        Dim f(1), i As Int64
        Dim s As String
        ReDim f(CInt(k))
        If c = "1" Then
            For i = 1 To CInt(k)
                s = s & Parity(i) & "."
            Next
            ThueMorse = s
        Else
            ThueMorse = Parity(k)
        End If
    End Function
    Private Function fibonachi(ByVal k As String, Optional ByVal c As Char = "0") As String
        Dim f(1), i As Int64
        Dim s As String
        ReDim f(CInt(k))
        f(0) = 0
        f(1) = 1
        f(2) = 1
        For i = 2 To CInt(k)
            f(i) = f(i - 1) + f(i - 2)
        Next i
        If c = "1" Then
            For i = 0 To CInt(k)
                s = s & Trim(Str(f(i))) & "."
            Next
            fibonachi = s
        Else
            fibonachi = Trim(Str(f(CInt(k))))
        End If
    End Function
    Private Function kor(ByVal n As Double, ByVal p As Double) As Double
        kor = n ^ (1 / p)
    End Function
    Private Function Prime(ByVal n As String) As Char
        Dim j As Int64
        Dim r As String
        For j = 2 To CInt(n) - 1
            r = remn(n, j)
            If r = "0" Or r = "-0" Then
                Prime = "0"
                Exit Function
            End If
        Next j
        Prime = "1"
    End Function

    Private Function Factorial(ByVal n As String) As String
        Dim j As Int64
        Dim s As String
        s = "1"
        For j = 1 To CLng(n)
            s = Mul(s, Trim(Str(j)))
        Next j
        Factorial = s
    End Function
    Private Function pow(ByVal x As String, ByVal n As String, Optional ByVal d As String = "10") As String
        Dim f As Boolean
        Dim i As Int64
        Dim z As Short
        Dim v, rp, s, n2, fr As String
        If x = "-1" Then
            If isodd(n) = "1" Then
                s = "-1"
            Else
                s = "1"
            End If
        Else

            s = "1"
            f = False
            If Strings.Left(n, 1) = "-" Then
                n = Mid(n, 2)
                f = True
            End If
            z = Strings.InStr(n, ".")
            If z = 0 Then
            Else
                z = Len(n) - z
            End If
            rp = "1" & stringd(z, "0")
            n = Val(Replace(n, ".", ""))
            v = GCD2(n, rp)
            n = div(n, v, 10)
            rp = div(rp, v, 10)
            'n = Int64.MaxValue
            For i = 1 To CLng(n)
                s = Mul(s, x, 0)
            Next i
            If z = 0 Then
            Else
                s = Rut(s, rp, d)
            End If
            If f = True Then
                s = div("1", s, d)
            End If
        End If
        pow = cut(s, d)
    End Function

    Private Function NjutSQRT(ByVal n As Double, ByVal p As Byte) As String
        Dim x As Double
        Dim f As Boolean
        Dim s As String
        Dim j As Byte
        f = False
        If n < 0 Then
            f = True
            n = -n
        End If
        x = 1
        For j = 1 To p
            x = 1 / 2 * (x + n / x)
        Next j
        If f = True Then
            s = x & "i"
        Else
            s = x & ""
        End If
        NjutSQRT = s
    End Function
    Private Function IsInteger(ByVal n As String) As Char
        If InStr(n, ",") > 0 Or InStr(n, ".") > 0 Then
            IsInteger = "0"
        Else
            IsInteger = "1"
        End If
    End Function
    Private Function Abso(ByVal a As String) As String
        Abso = Replace(a, "-", "")
    End Function
    Private Function ROOT(ByVal x As Double, ByVal n As Integer) As String
        Dim r As Double
        Dim f As Boolean
        Dim s As String
        f = False
        If x < 0 Then
            x = -x
            f = True
        End If
        r = x ^ (1 / n)
        If f = True Then
            s = r & "i"
        Else
            s = r & ""
        End If
        ROOT = Replace(s, ",", ".")
    End Function
    Private Function fromRoman(ByVal n As String) As String
        Dim i As Int32
        Dim c, cp, s As System.String
        i = Len(n)
        Do Until i < 1
            c = Mid(n, i, 1)
            If c = "I" Then
                If cp = "V" Or cp = "X" Then
                    s = subs(s, "1")
                Else
                    s = add(s, "1")
                End If
            End If
            If c = "V" Then
                s = add(s, "5")
            End If
            If c = "X" Then
                If cp = "L" Or cp = "C" Then
                    s = subs(s, "10")
                Else
                    s = add(s, "10")
                End If

            End If
            If c = "L" Then
                s = add(s, "50")
            End If
            If c = "C" Then
                If cp = "D" Or cp = "Ī" Then
                    s = subs(s, "100")
                Else
                    s = add(s, "100")
                End If
            End If
            If c = "D" Then
                s = add(s, "500")
            End If
            If c = "Ī" Then
                If cp = ChrW(12289) Or cp = ChrW(12288) Then
                    s = subs(s, "1000")
                Else
                    s = add(s, "1000")
                End If
            End If
            If c = ChrW(12289) Then
                s = add(s, "5000")
            End If
            If c = ChrW(12288) Then
                If cp = "|I|" Or cp = ChrW(313) Then
                    s = subs(s, "10000")
                Else
                    s = add(s, "10000")
                End If
            End If
            If c = ChrW(313) Then
                s = add(s, "50000")
            End If
            If c = "|" Then
                c = Mid(n, i - 2, 3)
                i = i - 2
                If c = "|I|" Then
                    If cp = "|V|" Or cp = "|X|" Then
                        s = subs(s, "100000")
                    Else
                        s = add(s, "100000")
                    End If
                End If
                If c = "|V|" Then
                    s = add(s, "500000")
                End If
                If c = "|X|" Then
                    If cp = "|L|" Or cp = "|C|" Then
                        s = subs(s, "1000000")
                    Else
                        s = add(s, "1000000")
                    End If
                End If
                If c = "|L|" Then
                    s = add(s, "5000000")
                End If
                If c = "|C|" Then
                    If cp = "|D|" Or cp = "|Ī|" Then
                        s = subs(s, "10000000")
                    Else
                        s = add(s, "10000000")
                    End If
                End If
                If c = "|D|" Then
                    s = add(s, "50000000")
                End If
                If c = "|Ī|" Then
                    If cp = "|" & ChrW(12289) & "|" Or cp = "|" & ChrW(12288) & "|" Then
                        s = subs(s, "100000000")
                    Else
                        s = add(s, "100000000")
                    End If
                End If
                If c = "|" & ChrW(12289) & "|" Then
                    s = add(s, "500000000")
                End If
                If c = "|" & ChrW(12288) & "|" Then
                    If cp = ChrW(12300) Then
                        s = subs(s, "1000000000")
                    Else
                        s = add(s, "1000000000")
                    End If
                End If
                If c = "|" & ChrW(12300) & "|" Then
                    s = add(s, "5000000000")
                End If
            End If
            cp = c
            i = i - 1
        Loop
        fromRoman = s
    End Function

    Private Function roman(ByVal n As String) As String
        Dim s As String
        Dim r As Byte
        If Comp(n, "1000000000") = "a<b" Then
            GoTo ne43
        Else
            r = CInt(idiv(n, "1000000000"))
            n = subs(n, Mul(r, "1000000000"))
        End If
        s = s & romanlg2(r, "|")
ne43:
        If Comp(n, "100000000") = "a<b" Then
            GoTo ne42
        Else
            r = CInt(idiv(n, "100000000"))
            n = subs(n, Mul(r, "100000000"))
        End If
        s = s & romansm2(r, "|")
ne42:
        If Comp(n, "10000000") = "a<b" Then
            GoTo ne1
        Else
            r = CInt(idiv(n, "10000000"))
            n = subs(n, Mul(r, "10000000"))
        End If
        s = s & romanmd(r, "|")
ne1:
        If Comp(n, "1000000") = "a<b" Then
            GoTo ne2
        Else
            r = CInt(idiv(n, "1000000"))
            n = subs(n, Mul(r, "1000000"))
        End If
        s = s & romanlg(r, "|")
ne2:
        If Comp(n, "100000") = "a<b" Then
            GoTo ne3
        Else
            r = CInt(idiv(n, "100000"))
            n = subs(n, Mul(r, "100000"))
        End If
        s = s & romansm(r, "|")
ne3:
        If Comp(n, "10000") = "a<b" Then
            GoTo ne41
        Else
            r = CInt(idiv(n, "10000"))
            n = subs(n, Mul(r, "10000"))
        End If
        s = s & romanlg2(r)
ne41:
        If Comp(n, "1000") = "a<b" Then
            GoTo ne4
        Else
            r = CInt(idiv(n, "1000"))
            n = subs(n, Mul(r, "1000"))
        End If
        s = s & romansm2(r)
ne4:
        If Comp(n, "100") = "a<b" Then
            GoTo ne7
        Else
            r = CInt(idiv(n, "100"))
            n = subs(n, Mul(r, "100"))
        End If
        s = s & romanmd(r)
ne7:
        If Comp(n, "10") = "a<b" Then
            GoTo ne9
        Else
            r = CInt(idiv(n, "10"))
            n = subs(n, Mul(r, "10"))
        End If
        s = s & romanlg(r)
ne9:
        If Comp(n, "1") = "a<b" Then
            GoTo nee
        Else
            r = CInt(idiv(n, "1"))
            n = subs(n, Mul(r, "1"))
        End If
        s = s & romansm(r)
nee:
        roman = s
    End Function
    Private Function romansm(ByVal n As Byte, Optional ByVal t As String = "") As String
        If n >= 1 And n < 4 Then
            romansm = stringd(n, t & "I" & t)
        End If
        If n = "4" Then
            romansm = t & "I" & t & t & "V" & t
        End If
        If n >= 5 And n < 9 Then
            romansm = t & "V" & t & stringd(n - 5, t & "I" & t)
        End If
        If n = "9" Then
            romansm = t & "I" & t & t & "X" & t
        End If
    End Function
    Private Function romanmd(ByVal n As Byte, Optional ByVal t As String = "") As String
        If n >= 1 And n < 4 Then
            romanmd = stringd(n, t & "C" & t)
        End If
        If n = "4" Then
            romanmd = t & "C" & t & t & "D" & t
        End If
        If n >= 5 And n < 9 Then
            romanmd = t & "D" & t & stringd(n - 5, t & "C" & t)
        End If
        If n = "9" Then
            romanmd = t & "C" & t & t & "Ī" & t
        End If
    End Function
    Private Function romanlg(ByVal n As Byte, Optional ByVal t As String = "") As String
        If n >= 1 And n < 4 Then
            romanlg = stringd(n, t & "X" & t)
        End If
        If n = "4" Then
            romanlg = t & "X" & t & t & "L" & t
        End If
        If n >= 5 And n < 9 Then
            romanlg = t & "L" & t & stringd(n - 5, t & "X" & t)
        End If
        If n = "9" Then
            romanlg = t & "X" & t & t & "C" & t
        End If
    End Function
    Private Function romanlg2(ByVal n As Byte, Optional ByVal t As String = "") As String
        If n >= 1 And n < 4 Then
            romanlg2 = stringd(n, t & ChrW(12288) & t)
        End If
        If n = "4" Then
            romanlg2 = t & ChrW(12288) & t & ChrW(12300) & t
        End If
        If n >= 5 And n < 9 Then
            romanlg2 = t & ChrW(12300) & t & stringd(n - 5, t & ChrW(12288) & t)
        End If
        If n = "9" Then
            romanlg2 = t & ChrW(12288) & t & t & "|I|" & t
        End If
    End Function


    Private Function romansm2(ByVal n As Byte, Optional ByVal t As String = "") As String
        If n >= 1 And n < 4 Then
            romansm2 = stringd(n, t & "Ī" & t)
        End If
        If n = "4" Then
            romansm2 = t & "Ī" & t & t & ChrW(12289) & t
        End If
        If n >= 5 And n < 9 Then
            romansm2 = t & ChrW(12289) & t & stringd(n - 5, t & "Ī" & t)
        End If
        If n = "9" Then
            romansm2 = t & "Ī" & t & t & ChrW(12288) & t
        End If
    End Function
    Private Function goldenratio() As String
        goldenratio = Mul(add(Rut(5, 2, 20), "1"), "0.5")
    End Function
    Private Function Parity(ByVal n As String) As String
        Dim c1, c0 As Int64
        n = convertB(n, "10", "2")
        c1 = CntChar(n, "1")
        If c1 Mod 2 = 0 Then
            Parity = "0"
        Else
            Parity = "1"
        End If
    End Function
    Private Function CntChar(ByVal n As String, ByVal h As Char) As String
        Dim i, k As Int64
        Dim c As Char
        For i = 1 To Len(n)
            c = Mid(n, i, 1)
            If c = h Then
                k = k + 1
            End If
        Next i
        CntChar = Trim(Str(k))
    End Function
    Private Function add(ByVal a As String, ByVal b As String) As String
        Dim f1, f2 As Byte
        Dim i, i1, m, n, m1, m2, ma As Int32
        Dim iz, ix As Int16
        Dim s As String
        Dim r3, f As Integer
        Dim c1, c2, ia, ib As Char

        m1 = 0
        m2 = 0
        ma = 0

        If Microsoft.VisualBasic.Left(a, 1) = "-" Then
            a = Replace(a, "-", "")
            ia = "-"
        End If
        If Microsoft.VisualBasic.Left(b, 1) = "-" Then
            b = Replace(b, "-", "")
            ib = "-"
        End If

        If (ib = "-" And ia = "-") Or (ib = Nothing And ia = Nothing) Then
            iz = 1
            ix = 1
        Else
            If Comp(a, b) = "a<b" Then
                ix = -1
                iz = 1
            Else
                ix = 1
                iz = -1
            End If
        End If


        For i = 1 To Len(a) Step 1
            c1 = Mid(a, Len(a) - i + 1, 1)
            If c1 = "." Then
                m1 = m1 + i - 1
            End If
        Next i

        For i = 1 To Len(b) Step 1
            c1 = Mid(b, Len(b) - i + 1, 1)
            If c1 = "." Then
                m2 = m2 + i - 1
            End If
        Next i
        If m1 > m2 Then
            ma = m1
            b = b & stringd(ma - m2, "0")
        Else
            ma = m2
            a = a & stringd(ma - m1, "0")
        End If
        a = Replace(a, ".", "")
        b = Replace(b, ".", "")
        If Len(a) > Len(b) Then
            f1 = 1
            f2 = 0
            m = Len(a)
            n = Len(b)
        Else
            f2 = 1
            f1 = 0
            m = Len(b)
            n = Len(a)
        End If
        If m = n Then
            m = m + 1
            f1 = 1
            f2 = 1
        End If
        f = 0
        For i1 = 1 To m
            If i1 > n And f2 = 1 Then
                c1 = "0"
            Else
                c1 = Mid(a, Len(a) - i1 + 1, 1)
            End If
            If i1 > n And f1 = 1 Then
                c2 = "0"
            Else
                c2 = Mid(b, Len(b) - i1 + 1, 1)
            End If
            r3 = Val(c1) * ix + Val(c2) * iz
            r3 = r3 + f
            If Len(Trim(Str(r3))) > 1 And Microsoft.VisualBasic.Left(Trim(Str(r3)), 1) <> "-" Then
                f = Val(Mid(Trim(Str(r3)), 1, Len(Trim(Str(r3))) - 1))
            Else
                If r3 < 0 Then
                    f = -1
                    r3 = 10 + r3
                Else
                    f = 0
                End If

            End If
            s = Val(Microsoft.VisualBasic.Right(Trim(Str(r3)), 1)) & s
            r3 = 0
        Next i1
        s = f & s
        If s = "" Then
            s = "0"
        End If
        If ma <> 0 Then
            s = Mid(s, 1, Len(s) - ma) & "." & Mid(s, Len(s) - ma + 1, ma)
        Else
            s = s
        End If
        s = shorts(s)

        If ib = "-" And ia = "-" Then
            s = "-" & s
        Else
            If ib = Nothing And ia = Nothing Then
            Else
                If ib = "-" Then
                    If iz = 1 Then
                        s = "-" & s
                    End If
                Else
                    If ix = 1 Then
                        s = "-" & s
                    End If
                End If
            End If
        End If

        add = s
    End Function

    Private Function checkzero(ByVal n As String) As String
        Dim s As String
        Dim a, f As Char
        Dim i As Int64
        n = Replace(n, ",", ".")
        f = "0"
        For i = 1 To Len(n)
            a = Mid(n, i, 1)
            If a = "0" Or a = "." Then
            Else
                f = "1"
                Exit For
            End If
        Next
        If f = "0" Then
            s = "0"
        Else
            s = n
        End If
        checkzero = s
    End Function

    Private Function truncateNull(ByVal n As String) As String
        Dim a, nt, nt2, sign As String
        Dim num, j, i As Integer
        nt = Replace(n, ",", ".")
        i = InStr(nt, ".")
        If i > 0 Then
            nt2 = Mid(nt, 1, i - 1)

        Else
            nt2 = nt
        End If
        For j = 1 To Len(nt2)
            a = Mid(nt, j, 1)
            If a = "0" Or (a = "-" And j = 1) Then
                If (a = "-" And j = 1) Then
                    sign = a
                End If
                num = j
            Else
                num = j
                Exit For
            End If
        Next
        If num = 0 Then
            nt = n
        Else
            nt = Mid(n, num)
        End If
        truncateNull = sign & nt
    End Function
    Private Function Comp(ByVal a As String, ByVal b As String) As String
        a = truncateNull(a)
        b = truncateNull(b)
        a = checkzero(a)
        b = checkzero(b)
        Dim ia, ib, ca, cb As Byte
        Dim i, m1, m2, ma As Int32
        Dim c1 As Char
        If Microsoft.VisualBasic.Left(a, 1) = "-" Then
            ia = 1
        End If
        If Microsoft.VisualBasic.Left(b, 1) = "-" Then
            ib = 1
        End If
        If a = b Then
            Comp = "a=b"
        End If
        a = Replace(a, "-", "")
        b = Replace(b, "-", "")
        For i = 1 To Len(a) Step 1
            c1 = Mid(a, Len(a) - i + 1, 1)
            If c1 = "." Then
                m1 = m1 + i - 1
            End If
        Next i
        For i = 1 To Len(b) Step 1
            c1 = Mid(b, Len(b) - i + 1, 1)
            If c1 = "." Then
                m2 = m2 + i - 1
            End If
        Next i
        If m1 > m2 Then
            ma = m1
            b = b & stringd(ma - m2, "0")
        Else
            ma = m2
            a = a & stringd(ma - m1, "0")
        End If
        a = Replace(a, ".", "")
        b = Replace(b, ".", "")

        If ((ia = 1 And ib = 1) Or (ia = 0 And ib = 0)) Then
            If Len(a) > Len(b) Then
                Comp = "a>b"
            End If
            If Len(a) < Len(b) Then
                Comp = "a<b"
            End If
            If Len(a) = Len(b) Then
                For i = 1 To Len(a)
                    ca = Val(Mid(a, i, 1))
                    cb = Val(Mid(b, i, 1))
                    If ca > cb Then
                        If (ia = 1 And ib = 1) Then
                            Comp = "a<b"
                        Else
                            Comp = "a>b"
                        End If
                        Exit For
                    End If
                    If ca < cb Then
                        If (ia = 1 And ib = 1) Then
                            Comp = "a>b"
                        Else
                            Comp = "a<b"
                        End If
                        Exit For
                    End If
                Next i
            End If
        Else
            If ia > ib Then
                Comp = "a<b"
            Else
                Comp = "a>b"
            End If
        End If
    End Function

    Private Function Mul(ByVal a As String, ByVal b As String, Optional ByVal c As Int16 = 0) As String
        Dim s, zn1, zn2 As String
        Dim c3, ia, ib, ma, mb As Char
        Dim i, i1, i2, m, r3, ist As Int64
        Dim chislo(1, 1)
        Dim r, c1, c2, r2, f As Byte
        m = 0
        ia = ""
        ib = ""
        ma = ""
        mb = ""
        a = Replace(a, ",", ".")
        b = Replace(b, ",", ".")
        If Microsoft.VisualBasic.Left(a, 1) = "-" Then
            a = Replace(a, "-", "")
            ia = "-"
        End If
        If Microsoft.VisualBasic.Left(b, 1) = "-" Then
            b = Replace(b, "-", "")
            ib = "-"
        End If

        If Microsoft.VisualBasic.Left(a, 1) = "i" Then
            a = Replace(a, "i", "")
            ma = "i"
        End If
        If Microsoft.VisualBasic.Left(b, 1) = "i" Then
            b = Replace(b, "i", "")
            mb = "i"
        End If

        For i = 1 To Len(a)
            c3 = Mid(a, Len(a) - i + 1, 1)
            If c3 = "." Then
                m = m + i - 1
            End If
        Next i
        For i = 1 To Len(b)
            c3 = Mid(b, Len(b) - i + 1, 1)
            If c3 = "." Then
                m = m + i - 1
            End If
        Next i
        ist = InStr(a, ".")
        If ist <> 0 Then
            If truncateNull(Mid(a, 1, ist - 1)) = "0" Then
                zn1 = "0"
            End If
        End If
        ist = InStr(b, ".")
        If ist <> 0 Then
            If truncateNull(Mid(b, 1, ist - 1)) = "0" Then
                zn2 = "0"
            End If
        End If
        a = Replace(a, ".", "")
        b = Replace(b, ".", "")
        ReDim chislo(Len(a) + Len(b), Len(b))


        For i1 = 1 To Len(b) Step 1
            For i2 = 1 To Len(a) Step 1
                c1 = Val(Mid(b, Len(b) - i1 + 1, 1))
                c2 = Val(Mid(a, Len(a) - i2 + 1, 1))
                r = c1 * c2 + f
                If Len(Trim(Str(r))) > 1 Then
                    f = Val(Mid(Trim(Str(r)), 1, 1))
                Else
                    f = 0
                End If
                r2 = Val(Microsoft.VisualBasic.Right(Trim(Str(r)), 1))
                chislo(i2 + i1 - 1, i1) = r2
            Next i2
            chislo(i2 + i1 - 1, i1) = f
            f = 0
        Next i1
        f = 0
        For i1 = 1 To Len(a) + Len(b)
            For i2 = 1 To Len(b)
                c1 = chislo(i1, i2)

                r3 = r3 + Val(c1)
            Next i2
            r3 = r3 + f
            If Len(Trim(Str(r3))) > 1 Then
                f = Val(Mid(Trim(Str(r3)), 1, Len(Trim(Str(r3))) - 1))
            Else
                f = 0
            End If
            s = Val(Microsoft.VisualBasic.Right(Trim(Str(r3)), 1)) & s
            r3 = 0
        Next i1

        If m <> 0 Then
            If c = 0 Then
                s = Mid(s, 1, Len(s) - m) & "." & Mid(s, Len(s) - m + 1)
            Else
                s = Mid(s, 1, Len(s) - m) & "." & Strings.Left(Mid(s, Len(s) - m + 1), Asc(c) - 5)
            End If

        Else
            s = s
        End If
        s = shorts(s)
        If (ia = "-" And ib = "-") Or (ia = Nothing And ib = Nothing) Then
        Else
            s = "-" & s
        End If

        Mul = s

    End Function

    Private Function Floor(ByVal n As String) As String
        Dim s As String
        n = shorts(n)
        If shorts(n) = shorts(ipart(n)) Then
            s = n
        Else
            If Comp(n, "0") = "a<b" Then
                s = add(ipart(n), "-1")
            Else
                s = ipart(n)
            End If
        End If
        If s = "-0" Then
            s = "0"
        End If
        Floor = s
    End Function

    Private Function Dividers(ByVal n As String) As String
        Dim r, st As String
        Dim j As Int64
        For j = 2 To n - 1
            r = remn(n, Trim(Str(j)))
            If r = "0" Or r = "-0" Then
                st = st & j & ";"
            End If
        Next j
        Dividers = st
    End Function
    Private Function hextodec(ByVal n As String) As String
        Dim c As Char
        Dim s As String
        Dim j, i As Int64
        n = Strings.Mid(n, 1, Len(n) - 1)
        For i = Len(n) To 1 Step -1
            c = Mid(n, i, 1)
            s = add(s, Mul(pow("16", j), cnva(c)))
            j = j + 1
        Next
        hextodec = s
    End Function
    Private Function perctonum(ByVal n As String) As String
        Dim c As Char
        Dim s As String
        Dim j, i As Int64
        n = Strings.Mid(n, 1, Len(n) - 1)
        perctonum = div(n, "100", 10)
    End Function
    Private Function octtodec(ByVal n As String) As String
        Dim c As Char
        Dim s As String
        Dim j, i As Int64
        n = Strings.Mid(n, 1, Len(n) - 1)
        For i = Len(n) To 1 Step -1
            c = Mid(n, i, 1)
            s = add(s, Mul(pow("8", j), c))
            j = j + 1
        Next
        octtodec = s
    End Function
    Private Function digiblock(ByVal n As String, ByVal ss As String, Optional ByVal b As Char = "10") As String
        Dim i As Integer
        Dim a, sn, s As String
        sn = "0"
        s = convertB(n, 10, b)
        For i = 1 To Len(s)
            a = Mid(s, i, 2)
            If a = ss Then
                sn = add(sn, 1)
            End If
        Next
        digiblock = sn
    End Function
    Private Function digisum(ByVal n As String, Optional ByVal b As Char = "10") As String
        Dim i As Integer
        Dim a, sn, s As String
        sn = "0"
        s = convertB(n, 10, b)
        For i = 1 To Len(s)
            a = Mid(s, i, 1)
            sn = add(sn, a)
        Next
        digisum = sn
    End Function
    Private Function digicount(ByVal n As String, Optional ByVal b As Char = "10") As String
        Dim i As Integer
        Dim a, s As String
        s = convertB(n, 10, b)

        digicount = Len(s)
    End Function
    Private Function digiroot(ByVal n As String) As String
        digiroot = add("1", remn(subs(n, 1), "9"))
    End Function
    Private Function eulerN(ByVal n As String, ByVal k As String) As String
        Dim b As Byte
        If CInt(n) < CInt(k) Then
            b = MsgBox("K>N!", MsgBoxStyle.AbortRetryIgnore, "Digit Processor Error")
            Exit Function
        End If
        Dim i, j As Int64
        Dim a(1, 1) As String
        Dim s As String
        ReDim a(CInt(n), CInt(n))
        For i = 1 To CInt(n)
            a(i, 1) = 1
            For j = 2 To CInt(i) - 1
                a(i, j) = add(Mul((i - j + 1), a(i - 1, j - 1)), Mul(j, a(i - 1, j)))
            Next
            a(i, i) = 1
        Next
        eulerN = a(CInt(n), CInt(k))
    End Function
    Private Function bintodec(ByVal n As String) As String
        Dim c As Char
        Dim s As String
        Dim j, i As Int64
        n = Strings.Mid(n, 1, Len(n) - 1)
        For i = Len(n) To 1 Step -1
            c = Mid(n, i, 1)
            s = add(s, Mul(pow("2", j), c))
            j = j + 1
        Next
        bintodec = s
    End Function
    Private Function cnva(ByVal n As Char) As String
        Dim s As Byte
        n = UCase(n)
        If Asc(n) > 64 Then
            s = Asc(n) - 55
        Else
            s = CByte(Val(n))
        End If
        cnva = s
    End Function
    Private Function compare(ByVal s As String) As Boolean
        Dim z As Int64
        Dim op As Byte
        Dim exp1, exp2 As String
        z = InStr(s, ">>")
        If z <> 0 Then
            op = 1
            GoTo begin
        End If
        z = InStr(s, "<<")
        If z <> 0 Then
            op = 2
            GoTo begin
        End If
        z = InStr(s, "==")
        If z <> 0 Then
            op = 3
            GoTo begin
        End If
        z = InStr(s, "!=")
        If z <> 0 Then
            op = 4
            GoTo begin
        End If
        z = InStr(s, "<>")
        If z <> 0 Then
            op = 4
            GoTo begin
        End If
        z = InStr(s, "><")
        If z <> 0 Then
            op = 4
            GoTo begin
        End If
        z = InStr(s, "=>")
        If z <> 0 Then
            op = 6
            GoTo begin
        End If
        z = InStr(s, ">=")
        If z <> 0 Then
            op = 6
            GoTo begin
        End If
        z = InStr(s, "<=")
        If z <> 0 Then
            op = 5
            GoTo begin
        End If
        z = InStr(s, "=<")
        If z <> 0 Then
            op = 5
            GoTo begin
        End If
        z = InStr(s, "!<")
        If z <> 0 Then
            op = 6
            GoTo begin
        End If
        z = InStr(s, "!>")
        If z <> 0 Then
            op = 5
            GoTo begin
        End If
begin:
        exp1 = Maincnt(Mid(s, 1, z - 1))
        exp2 = Maincnt(Mid(s, z + 2))
        If op = 2 Then
            If Comp(exp1, exp2) = "a<b" Then
                compare = True
            Else
                compare = False
            End If
            Exit Function
        End If
        If op = 1 Then
            If Comp(exp1, exp2) = "a>b" Then
                compare = True
            Else
                compare = False
            End If
            Exit Function
        End If
        If op = 3 Then
            If Comp(exp1, exp2) = "a=b" Then
                compare = True
            Else
                compare = False
            End If
            Exit Function
        End If
        If op = 4 Then
            If Comp(exp1, exp2) = "a=b" Then
                compare = False
            Else
                compare = True
            End If
            Exit Function
        End If
        If op = 5 Then
            If Comp(exp1, exp2) = "a>b" Then
                compare = False
            Else
                compare = True
            End If
            Exit Function
        End If
        If op = 6 Then
            If Comp(exp1, exp2) = "a<b" Then
                compare = False
            Else
                compare = True
            End If
            Exit Function
        End If
    End Function
    Private Function acnv(ByVal n As Int64) As String
        Dim s As Char
        n = UCase(n)
        If n > 9 Then
            s = Chr(n + 55)
        Else
            s = n & ""
        End If
        acnv = s
    End Function
    Private Function Maya_acnv(ByVal n As Int64) As String
        Dim s As Char
        n = UCase(n)
        If n <> "0" Then
            s = ChrW(n + 12316)
        Else
            s = ChrW(&H3036)
        End If
        Maya_acnv = s
    End Function
    Private Function Ceiling(ByVal n As String) As String
        Dim s As String
        n = shorts(n)
        n = shorts(n)
        If shorts(n) = shorts(ipart(n)) Then
            s = n
        Else
            If Comp(n, "0") = "a<b" Then
                s = ipart(n)
            Else
                s = add(ipart(n), "1")
            End If
        End If
        If s = "-0" Then
            s = "0"
        End If
        Ceiling = s
    End Function
    Private Function Congruent(ByVal b As String, ByVal c As String, ByVal m As String) As Char
        Dim s As String
        s = remn(subs(b, c), m)
        If s = "0" Or s = "-0" Then
            Congruent = "1"
        Else
            Congruent = "0"
        End If
    End Function
    Private Function Frac1(ByVal n As String) As String
        If n >= 0 Then
            Frac1 = subs(n, Floor(n))
        Else
            Frac1 = subs(n, subs(Floor(n), 1))
        End If
    End Function
    Private Function Frac2(ByVal n As String) As String
        Frac2 = subs(n, Floor(n))
    End Function
    Private Function IntPart(ByVal n As String) As String
        If Comp(n, 0) = "a<b" Then
            IntPart = Ceiling(n)
        Else
            IntPart = Floor(n)
        End If
    End Function

    Private Function Rut(ByVal x As String, ByVal n As String, ByVal t As Integer) As String
        Dim zu, l, z, e As Int64
        Dim x2, xe, sp, isi, p, sr, r, ra, sm, s, bef As String
        Dim j As Short
        Dim i As Byte
        If Comp(x, 0) = "a>b" Then
        Else
            If truncateNull(x) = "0" Then
                s = 0
                GoTo ennt
            Else
                If isodd(n) = "0" Then
                    s = "i"
                    x = Mid(x, 2)
                End If
            End If
        End If
        '
        e = InStr(1, x, ".")
        If e > 0 Then
            x2 = Strings.Left(x, e - 1)
        Else
            x2 = x
        End If

        'zu = Len(x) - e
        '
        '       If e = 0 Then
        '      zu = 0
        '     Else
        '        zu = zu ^ n
        '    End If
        If x2 = 0 Then
            l = 1
        Else
            l = Len(subs(x2, "1"))
        End If
        xe = (l \ 2 + l Mod 2)
        p = l - n '- 1
        If p < 0 Then
            p = 0
        End If
fors:
        sp = 0
        For j = 1 To t
            For i = 1 To 10
                sr = s & i
                If z = 0 Then
                    sr = sr
                Else
                    bef = Mid(sr, 1, Len(sr) - z)
                    If bef = "" Then
                        bef = "0"
                    End If
                    sr = bef & "." & Mid(sr, Len(sr) - z + 1)
                End If
                ra = sr & stringd(p, "0")
                r = pow(ra, n, t)
                If r <> x Then
                    'ra = r & String(l, "0")
                    If Comp(r, x) = "a>b" Then
                        s = s & sp
                        Exit For
                    Else
                        sp = i
                    End If
                Else
                    s = s & i
                    z = z + 1
                    GoTo ennt
                End If
            Next i
            If p = 0 Then
                z = z + 1
            Else
                p = p - 1
            End If
            sp = 0
        Next j
ennt:
        If z = 0 Then
            Rut = s
        Else
            Rut = Mid(s, 1, Len(s) - z + 1) & "." & Mid(s, Len(s) - (z - 2))
        End If
    End Function
    Private Function Logxb(ByVal x As String, ByVal b As String, Optional ByVal l As String = "10") As String
        Dim s As String

        If Comp(x, 0) <> "a>b" Then
            s = 0
        Else
            If Comp(b, 0) <> "a>b" Then
                s = 0
            Else
                s = div(Logn(x), Logn(b), l)
            End If
        End If
        Logxb = s
    End Function
    Private Function Lognm(ByVal x As String) As String
        Dim s, x2, tnp1, zdz, p As String
        Dim i As Int64
        zdz = div(subs(x, "1"), add(x, "1"), 20)
        For i = 0 To 15
            tnp1 = add(Mul(2, Trim(Str(i))), 1)
            p = pow(zdz, tnp1, 20)
            s = add(s, Mul(div(1, tnp1, 20), p))
        Next i
        Lognm = Mul(s, "2")
    End Function
    Private Function Logn(ByVal x As String) As String
        Dim s, x2, x3, xp As String
        Dim i As Int64
        i = 0
        x2 = "0"
        If Comp(x, "2") = "a<b" Then
            Logn = Lognm(x)
        Else
            Do Until Comp(x2, x) = "a>b"
                i = i + 1
                xp = x2
                x2 = pow(2, i)
            Loop
            x3 = div(x, xp, 10)
            Logn = add(Lognm(x3), Mul(Trim(Str(i - 1)), ln2()))
        End If
    End Function
    Private Function ln2() As String
        ln2 = "0.693147180559945309417232121458176568075500134360255254120680009493393621969694715605863326996418687"
    End Function
    Private Function Logn2(ByVal x As String) As String
        Dim s, x2 As String
        Dim i As Int64
        For i = 1 To Val(x)
            s = add(s, div(1, Mul(Trim(Str(i)), pow(2, Trim(Str(i)))), Val(x)))
        Next i
        Logn2 = s
    End Function
    Private Function Round(ByVal n As Double) As Long
        Round = CInt(n)
    End Function
    Private Function RoundMath(ByVal n As String, ByVal dn As String) As String
        Dim s, nn, sy As String
        Dim z, rdn As Int64
        n = Replace(n, ",", ".")
        z = InStr(n, ".")
        If z = 0 Then
            s = n
            GoTo endd
        End If
        rdn = Len(n) - z
        s = Mid(n, 1, z + dn)
        sy = Mid(n, z + dn + 1, 1)
        If Comp(sy, "5") = "a>b" Then
            nn = "0." & stringd(dn - 1, "0") & "1"
            s = add(s, nn)
        Else
        End If
endd:
        RoundMath = s
    End Function
    Private Function RCINT(ByVal n As String) As String
        Dim a, a1, a2 As String
        If Comp(n, 0) = "a>b" Then
            a1 = Ceiling(n)
            a2 = Floor(n)
        End If
        If Comp(n, 0) = "a<b" Then
            a1 = Floor(n)
            a2 = Ceiling(n)
        End If
        If Comp(Abso(subs(n, a1)), Abso(subs(n, a2))) = "a>b" Then
            a = a2
        Else
            a = a1
        End If
        RCINT = a
    End Function
    Private Function Exponent() As String
        Exponent = Trim(Str(Math.Exp(1)))
    End Function
    Private Function PermutNum(ByVal x As String, ByVal n As String) As String
        PermutNum = pow(x, n)
    End Function

    Private Function CatTriangle(ByVal n As String) As String
        Dim s As String
        Dim k As Int64
        For k = 0 To CInt(n)
            s = s & (Mul(Factorial(n + k), Trim(Str((n - k + 1))))) / (Mul(Factorial(k), Factorial(n + 1))) & "."
        Next k
        CatTriangle = Microsoft.VisualBasic.Left(s, Len(s) - 1)
    End Function

    Private Function BinomCoef(ByVal n As String, ByVal k As String) As String
        If Comp(k, "0") = "a<b" Then
            BinomCoef = 0
            Exit Function
        End If
        If Comp(k, n) = "a>b" Then
            BinomCoef = 0
            Exit Function
        End If
        BinomCoef = div(Factorial(n), (Mul(Factorial(subs(n, k)), Factorial(k))), 10)
    End Function
    Private Function CentBiCoef(ByVal n As Integer) As Long
        CentBiCoef = BinomCoef(2 * n, n)
    End Function
    Private Function Factorial2(ByVal n As String) As String
        Dim r, s As String
        Dim e As Byte
        Dim i As Int64
        s = "1"
        r = remn(n, 2)
        If r = 0 Then
            e = 2
        Else
            e = 1
        End If
        If n = 0 Or n = -1 Then
            s = "1"
        Else
            For i = n To e Step -2
                s = Mul(s, Trim(Str(i)))
            Next i
        End If
        Factorial2 = s
    End Function
    Private Function MultiFactorial(ByVal n As String, ByVal k As String) As String
        Dim i As Int64
        Dim e, r, s As String
        s = "1"
        r = remn(n, k)
        If r = "0" Or r = "-0" Then
            e = k
        Else
            e = r
        End If
        For i = CInt(n) To CInt(e) Step -k
            s = Mul(s, Trim(Str(i)))
        Next i
        MultiFactorial = s
    End Function
    Private Function PasTriangle(ByVal n As Integer) As String
        Dim s As String
        Dim k As Int64
        For k = 0 To n
            s = s & BinomCoef(n, k) & "."
        Next k
        PasTriangle = Microsoft.VisualBasic.Left(s, Len(s) - 1)
    End Function

    Private Function count_elem(ByVal x As String, ByVal c As Char) As Int32
        Dim r, a As String
        Dim i, k, b, w As Int64
        'присвоить r - введенную строку, а в х сохранить исходную (на всякий случай)
        r = x
        For i = 1 To Len(r)
            ' считываем очередной символ из строки
            a = Mid(r, i, 1)
            'если попалась открытая скобка, то плюс один к счетчику скобок
            If a = "(" Or a = "[" Then
                w = w + 1
            End If
            'если попалась закрытая скобка, то минус один к счетчику скобок
            'цель этих счетчиков, учесть знаки умножения только в первом уровене выражения,
            'дальше дело делает рекурсия функций
            If a = ")" Or a = "]" Then
                w = w - 1
            End If
            'Если счетчик 0, т.е. первый уровень, то учитываю знак умножения 
            'и пишу в массив то что в скобках - размер массива 500, 
            'т.е. макс. 500 множителей на одном уровне ассоциативности
            If a = c And w = 0 Then
                k = k + 1
                b = i + 1
            End If
            'При появлении знака сложения или вычитания выйти из цикла
            If (a = "+" Or a = "-") And i <> 1 And w = 0 Then
                Exit For
            End If
        Next i
        'добиваю последний множитель 
        k = k + 1
        count_elem = k
    End Function
    Private Function Deriv1(ByVal x As String) As String
        Dim s, r, a, h, h0 As String
        Dim i, k, b, w, e, f, b2, j As Int64
        Dim da(0) As String
        'присвоить r - введенную строку, а в х сохранить исходную (на всякий случай)
        r = x
        'Заменить пустышки (х) 
        r = Replace(r, "(x)", "x")
        'Выразить деление через умножение
        r = Replace(r, "/x", "*x^-1") '-для простых
        r = replace_div(r, 1) ' - для сложных конструкций со скобками
        'заменить где можно двойную степень минус 1 и степень n числа на степень минус n
        r = Replace(r, "^-1^", "^-")
        'заменить степень из двух минусов
        r = Replace(r, "^--", "^")
        'заменить первую степень
        r = Replace(r, "^10", "$0")
        r = Replace(r, "^11", "$1")
        r = Replace(r, "^12", "$2")
        r = Replace(r, "^13", "$3")
        r = Replace(r, "^14", "$4")
        r = Replace(r, "^15", "$5")
        r = Replace(r, "^16", "$6")
        r = Replace(r, "^17", "$7")
        r = Replace(r, "^18", "$8")
        r = Replace(r, "^19", "$9")
        r = Replace(r, "^1", "")
        r = Replace(r, "$0", "^10")
        r = Replace(r, "$1", "^11")
        r = Replace(r, "$2", "^12")
        r = Replace(r, "$3", "^13")
        r = Replace(r, "$4", "^14")
        r = Replace(r, "$5", "^15")
        r = Replace(r, "$6", "^16")
        r = Replace(r, "$7", "^17")
        r = Replace(r, "$8", "^18")
        r = Replace(r, "$9", "^19")


        ' предустановки для цикла биения строки на множители
        k = 0
        w = 0
        f = 0
        b = 1
        'если в строке есть знаки умножения то делать этот цикл
        If InStr(r, "*") > 0 Then
            'Вызов функции подсчитывающей количество множителей на выбранном уровне ассоциативности
            'На деле функции повторяет нижеследующий цикл, только не производит никаких сохранений 
            'в массив, а нужна для осуществления операции redim, то есть для изменения размерности
            'массива для сохранения множителей.
            k = count_elem(r, "*") + 1
            'Делаем изменение размерности массива в соответствии с надобностью
            ReDim da(k)
            'Не забыть очистить использованную переменную k!!!
            If k < 3 Then
                GoTo entt
            End If
            k = 0
            'пробегаем всю длину строки и проверяем на наличие знаков умножения
            For i = 1 To Len(r)
                ' считываем очередной символ из строки
                a = Mid(r, i, 1)
                'если попалась открытая скобка, то плюс один к счетчику скобок
                If a = "(" Or a = "[" Then
                    w = w + 1
                End If
                'если попалась закрытая скобка, то минус один к счетчику скобок
                'цель этих счетчиков, учесть знаки умножения только в первом уровене выражения,
                'дальше дело делает рекурсия функций
                If a = ")" Or a = "]" Then
                    w = w - 1
                    If w = 0 Then
                        e = i + 1
                    End If
                End If
                'Если счетчик 0, т.е. первый уровень, то учитываю знак умножения 
                'и пишу в массив то что в скобках - размер массива 500, 
                'т.е. макс. 500 множителей на одном уровне ассоциативности
                If a = "*" And w = 0 Then
                    e = i
                    f = 1
                    k = k + 1
                    da(k) = Mid(r, b, e - b)
                    b = i + 1
                End If
                'При появлении знака сложения или вычитания выйти из цикла
                If (a = "+" Or a = "-") And i <> 1 And w = 0 Then
                    Exit For
                End If
            Next i
            e = i
            'добиваю последний множитель 
            If e = Len(r) + 1 And k = 0 And b = 1 Then
                k = k + 1
                da(k) = Mid(r, b + 1, e - b - 2)
                f = 2
            Else
                k = k + 1
                da(k) = Mid(r, b, e - b)
            End If
            'Записываю в h возможный хвостик (остаток)
            h0 = Mid(r, e, 1)
            h = Mid(r, e + 1)
            'в k сохранено количество множителей на данном уровне ассоциативности
            If f = 1 Then
                'согласно формуле 1.2 для производной первого порядка, от произведения из n множителей
                'одной ассоциативности, выполняю вычисления производных с рекурсией, всех уровней 
                'ассоциативности, т.е. пока не получаю производную простой функции.
                For i = 1 To k
                    If InStr(da(i), "x") = 0 Then
                        If i = 1 Then
                        Else
                            s = s & "0"
                        End If

                    Else
                        For j = 1 To i - 1
                            s = s & da(j) & "*"
                        Next j


                        s = s & Deriv1(da(i))
                        For j = i + 1 To k
                            s = s & "*" & da(j)
                        Next j
                        If i < k Then
                            s = s & "+"
                        End If
                    End If
                Next i
                'Данная формула 1.2 записана в моей записной черной книжке.
                'А теперь вспоминаем про хвостик, и прибавляем вычисленный рекурсивно хвостик
                'к полученной производной произведения.
                If h <> "" Then
                    s = s & h0 & Deriv1(h)
                End If
                '(x+1)*(x-1)+(x^2+1)*(x^2-1)
            Else
                If f = 0 Then
                    s = s & Deriv1(da(k)) & h0 & Deriv1(h)
                Else
                    s = s & Deriv1(da(k))
                End If
            End If
        Else
entt:
            'Вызвать производную сложной функции
            r = DerivPower(r)
            s = r
        End If
        Deriv1 = s
    End Function
    Private Function DerivPower(ByVal x As String) As String
        Dim s, r, a, h, h1, h0, p As String
        Dim da(0) As String
        Dim f, l, i, k, b, w, e, b2, j As Int64
        'присвоить r - введенную строку, а в х сохранить исходную (на всякий случай)
        r = x
        ' предустановки для цикла биения строки на множители
        k = 0
        w = 0
        b = 1
        f = 0
        If InStr(r, "^") > 0 Then
            'Вызов функции подсчитывающей количество множителей на выбранном уровне ассоциативности
            'На деле функции повторяет нижеследующий цикл, только не производит никаких сохранений 
            'в массив, а нужна для осуществления операции redim, то есть для изменения размерности
            'массива для сохранения множителей.
            k = count_elem(r, "^") + 1
            'Делаем изменение размерности массива в соответствии с надобностью
            ReDim da(k)
            'Не забыть очистить использованную переменную k!!!
            k = 0
            'пробегаем всю длину строки и проверяем на наличие знаков степени
            For i = 1 To Len(r)
                ' считываем очередной символ из строки
                a = Mid(r, i, 1)
                'если попалась открытая скобка, то плюс один к счетчику скобок
                If a = "(" Or a = "[" Then
                    w = w + 1
                End If
                'если попалась закрытая скобка, то минус один к счетчику скобок
                'цель этих счетчиков, учесть знаки степени только в первом уровене выражения,
                'дальше дело делает рекурсия функций
                If a = ")" Or a = "]" Then
                    w = w - 1
                    If w = 0 Then
                        e = i + 1
                    End If
                End If
                'Если счетчик 0, т.е. первый уровень, то учитываю знак степени 
                'и пишу в массив то что в скобках - размер массива 500, 
                'т.е. макс. 500 возведений в степень на одном уровне ассоциативности степени
                If a = "^" And w = 0 Then
                    e = i
                    f = 1
                    k = k + 1
                    da(k) = Mid(r, b, e - b)
                    b = i + 1
                End If
                If (a = "+" Or a = "-") And w = 0 Then
                    Exit For
                End If
            Next i
            'добиваю последнюю степень 
            e = i
            If e = Len(r) + 1 And k = 0 And b = 1 Then
                k = k + 1
                da(k) = Mid(r, b + 1, e - b - 2)
                a = Mid(r, 1, 1)
                If (Asc(a) > 97 And Asc(a) < 122) And a <> "x" Then
                    GoTo funct
                End If
                'da(k) = Mid(r, b, e - b)
                f = 2
            Else
                    k = k + 1
                    da(k) = Mid(r, b, e - b)
                End If
                'Записываю в h возможный хвостик (остаток)
                h1 = Mid(r, 1, e - 1)
                h0 = Mid(r, e, 1)
                h = Mid(r, e + 1)
                'в к сохранено количество степеней на данном уровне ассоциативности

                If f = 1 Then
                    'согласно формуле 2.5 для производной первого порядка, от степени из n элементов
                    'одной ассоциативности, выполняю вычисления производных с рекурсией, всех уровней 
                    'ассоциативности, т.е. пока не получаю производную простой функции.
                    For i = 0 To k - 2
                        s = s & stringd(i, "ln[") & h1 & stringd(i, "]") & "*"
                    Next i
                    If Deriv1(da(k)) = "0" Then
                        s = s & "(" & Deriv1(da(k - 1)) & "/" & da(k - 1) & "*" & da(k)
                    Else
                        If Deriv1(da(k - 1)) = "0" Then
                            s = s & "(ln[" & da(k - 1) & "]*" & Deriv1(da(k))
                        Else
                            s = s & "(ln[" & da(k - 1) & "]*" & Deriv1(da(k)) & "+" & Deriv1(da(k - 1)) & "/" & da(k - 1) & "*" & da(k)
                        End If

                    End If
                    For i = 2 To k - 1
                        j = k - i
                        For l = 0 To i - 1
                            p = p & "*" & stringd(l, "ln[") & da(j) & stringd(l, "]")
                        Next l
                        s = s & "+" & Deriv1(da(j)) & "/(" & Mid(p, 2) & ")"
                    Next i
                    s = s & ")"
                    'Данная формула 2.5 записана в моей записной черной книжке.
                    'А теперь вспоминаем про хвостик, и прибавляем вычисленный рекурсивно хвостик
                    'к полученной производной произведения.
                    If h <> "" Then
                        s = s & h0 & Deriv1(h)
                    End If
                    '(x+1)*(x-1)+(x^2+1)*(x^2-1)
                Else
                    If f = 0 Then
                        s = s & DerivBase(da(k)) & h0 & Deriv1(h)
                    Else
                        s = s & Deriv1(da(k))
                    End If
                End If
        Else
funct:
            r = DerivLn(r)
            s = r
        End If
        s = Deriv_clear(s)
        DerivPower = s
    End Function
    Private Function Deriv_clear(ByVal x As String) As String
        Dim s, r1, r2 As String
        Dim i, i1, i2, z0, z1, z2, ze As Int16
        'присвоить s - введенную строку, а в х сохранить исходную (на всякий случай)
        s = x
        Dim zn(4) As Char
        zn(1) = "+"
        zn(2) = "-"
        zn(3) = ")"
        zn(4) = "("
        'Здесь удаляем лишнее умножение на 1
        For i = 1 To 4
            s = Replace(s, "*1" + zn(i), zn(i))
            s = Replace(s, zn(i) + "1*", zn(i))
        Next
        'Здесь превращаем x деленное на x в 1
        For i1 = 1 To 4
            For i2 = 1 To 4
                s = Replace(s, zn(i1) + "1/x*x" + zn(i2), zn(i1) + "1" + zn(i2))
            Next
        Next

        Deriv_clear = s
    End Function
    Private Function DerivLn(ByVal x As String) As String
        Dim s As String
        s = x
        Dim z0, zE As Int32
        z0 = InStr(x, "ln[")
        zE = Len(x)
        If z0 = 0 Then
            s = DerivLog(s)
        Else
            z0 = z0 + 3
            s = Mid(x, z0, zE - z0)
            s = "(" + Derivation(s, 1) + ")/(" + s + ")"
        End If

        DerivLn = s
    End Function
    Private Function DerivLog(ByVal x As String) As String
        Dim s, s1, s2 As String
        s = x
        Dim z0, z1, zE As Int32
        z0 = InStr(x, "log[")
        z1 = InStr(x, ",")
        zE = Len(x)
        If z0 = 0 Then
            s = DerivSin(s)
        Else
            z0 = z0 + 3
            s1 = Mid(x, z0 + 1, z1 - z0 - 1)
            s2 = Mid(x, z1 + 1, zE - z1 - 1)
            s = Derivation("ln[" & s2 & "]/ln[" & s1 & "]", 1)
        End If

        DerivLog = s
    End Function
    Private Function DerivSin(ByVal x As String) As String
        Dim s, s1, s2 As String
        s = x
        Dim z0, z1, zE As Int32
        z0 = InStr(x, "sin[")
        zE = Len(x)
        If z0 = 0 Then
            s = DerivCos(s)
        Else
            z0 = z0 + 3
            s = Mid(x, z0, zE - z0)
            s = "(" + Derivation(s, 1) + ")*cos[" + s + "]"
        End If

        DerivSin = s
    End Function

    Private Function DerivCos(ByVal x As String) As String
        Dim s, s1, s2 As String
        s = x
        Dim z0, z1, zE As Int32
        z0 = InStr(x, "cos[")
        zE = Len(x)
        If z0 = 0 Then
            s = DerivCot(s)
        Else
            z0 = z0 + 3
            s = Mid(x, z0, zE - z0)
            s = "(" + Derivation(s, 1) + ")*(-1)*sin[" + s + "]"
        End If

        DerivCos = s
    End Function

    Private Function DerivCot(ByVal x As String) As String
        Dim s, s1, s2 As String
        s = x
        Dim z0, z1, zE As Int32
        z0 = InStr(x, "cot[")
        zE = Len(x)
        If z0 = 0 Then
            s = DerivTan(s)
        Else
            z0 = z0 + 3
            s = Mid(x, z0, zE - z0)
            s = "(" + Derivation(s, 1) + ")*(-1)/(sin[" + s + "])^2"
        End If

        DerivCot = s
    End Function

    Private Function DerivTan(ByVal x As String) As String
        Dim s, s1, s2 As String
        s = x
        Dim z0, z1, zE As Int32
        z0 = InStr(x, "tan[")
        zE = Len(x)
        If z0 = 0 Then
            s = DerivASin(s)
        Else
            z0 = z0 + 3
            s = Mid(x, z0, zE - z0)
            s = "(" + Derivation(s, 1) + ")*(-1)/(cos[" + s + "])^2"
        End If

        DerivTan = s
    End Function

    Private Function DerivASin(ByVal x As String) As String
        Dim s, s1, s2 As String
        s = x
        Dim z0, z1, zE As Int32
        z0 = InStr(x, "asin[")
        zE = Len(x)
        If z0 = 0 Then
            s = DerivACos(s)
        Else
            z0 = z0 + 3
            s = Mid(x, z0, zE - z0)
            s = "(" + Derivation(s, 1) + ")/(1-(" + s + ")^2)^(1/2)"
        End If

        DerivASin = s
    End Function

    Private Function DerivACos(ByVal x As String) As String
        Dim s, s1, s2 As String
        s = x
        Dim z0, z1, zE As Int32
        z0 = InStr(x, "acos[")
        zE = Len(x)
        If z0 = 0 Then
            s = DerivACot(s)
        Else
            z0 = z0 + 3
            s = Mid(x, z0, zE - z0)
            s = "(" + Derivation(s, 1) + ")*(-1)/(1-(" + s + ")^2)^(1/2)"
        End If

        DerivACos = s
    End Function

    Private Function DerivACot(ByVal x As String) As String
        Dim s, s1, s2 As String
        s = x
        Dim z0, z1, zE As Int32
        z0 = InStr(x, "acot[")
        zE = Len(x)
        If z0 = 0 Then
            s = DerivATan(s)
        Else
            z0 = z0 + 3
            s = Mid(x, z0, zE - z0)
            s = "(" + Derivation(s, 1) + ")*(-1)/(1+(" + s + ")^2)"
        End If

        DerivACot = s
    End Function

    Private Function DerivATan(ByVal x As String) As String
        Dim s, s1, s2 As String
        s = x
        Dim z0, z1, zE As Int32
        z0 = InStr(x, "atan[")
        zE = Len(x)
        If z0 = 0 Then
            s = DerivBase(s)
        Else
            z0 = z0 + 3
            s = Mid(x, z0, zE - z0)
            s = "(" + Derivation(s, 1) + ")/(1+(" + s + ")^2)"
        End If

        DerivATan = s
    End Function
    Private Function DerivBase(ByVal x As String) As String
        If InStr(x, "x") = 0 Then
            x = "0"
        Else
            If x = "x-x" Then
                x = "0"
            Else
                If InStr(x, "-x") <> 0 Then
                    x = "-1"
                Else
                    x = "1"
                End If
            End If


        End If
ennnd:
        DerivBase = x
    End Function
    Private Function replace_div(ByVal x As String, ByVal ind As Int16) As String
        Dim r, a, a1 As String
        Dim i, b, e As Int64
        'присвоить r - введенную строку, а в х сохранить исходную (на всякий случай)
        r = x
        'гонять рекурсию до тех пор пока все деления не заменены на умножения в -1 степени
        'на всех уровнях ассоциативности выражения
        If InStr(r, "/") > 0 Then
replay:
            ' предустановки для цикла биения строки на множители
            'Очень ВАЖНО сбросить данные переменные, т.к. см replay
            b = 0
            e = 0
            'пробегаем всю длину строки 
            For i = 1 To Len(r)
                ' считываем очередной символ из строки
                a = Mid(r, i, 1)
                'если попалось деление отмечаем начало и продолжаем цикл
                If a = "/" Then
                    b = i
                End If
                'если попалась закрытая скобка, и до этого попалось хоть одно деление, то меняем
                'рекурсивно все деления на умножение в -1 степени
                If (a = ")" Or a = "]") And b <> 0 Then
                    e = i
                    r = Mid(r, 1, b - 1) & replace_div(Mid(r, b + 1, e - b), 0) & Mid(r, e + 1)
                    'заменили выражение, перестроили, и надо гнать цикл заново, поехали на replay
                    GoTo replay
                End If
            Next i
        Else
            'если в функцию передано выражение без знаков деления, то приписываем ему "крылышки",
            'там есть индикатор чтобы не казнить невинных, но если виноват, то крылышки обз.прицепим
            If ind = 0 Then
                r = "*" & r & "^(-1)"
            End If
        End If
        replace_div = r

    End Function
    Private Function Derivation(ByVal x As String, ByVal n As String) As String
        Dim r As String
        Dim i As Int16
        x = LCase(x) '- чтобы однозначно распознять х
        'если производная нулевого порядка, то это есть сама введенная функция
        If Comp(n, "0") = "a=b" Then
            r = x
            GoTo endr
        End If
        'производные высших порядков сюда
        'Повторять этот цикл столько раз пока не вычислим произв. указ. порядка или функция не обратиться в 0
        For i = 1 To CInt(n)
            'ссылка на базовую процедуру нахождения производной 1 го порядка
            r = Deriv1(x)
            r = Deriv_clear(r)
            r = Replace(r, "ln[x]*0", "0")
            r = Replace(r, "+0", "")
            r = Replace(r, "-0", "")

            r = Replace(r, "10-", "$1")
            r = Replace(r, "20-", "$2")
            r = Replace(r, "30-", "$3")
            r = Replace(r, "40-", "$4")
            r = Replace(r, "50-", "$5")
            r = Replace(r, "60-", "$6")
            r = Replace(r, "70-", "$7")
            r = Replace(r, "80-", "$8")
            r = Replace(r, "90-", "$9")
            r = Replace(r, "00-", "$0")

            r = Replace(r, "0-", "-")

            r = Replace(r, "$1", "10-")
            r = Replace(r, "$2", "20-")
            r = Replace(r, "$3", "30-")
            r = Replace(r, "$4", "40-")
            r = Replace(r, "$5", "50-")
            r = Replace(r, "$6", "60-")
            r = Replace(r, "$7", "70-")
            r = Replace(r, "$8", "80-")
            r = Replace(r, "$9", "90-")
            r = Replace(r, "$0", "00-")

            r = Replace(r, "10+", "$1")
            r = Replace(r, "20+", "$2")
            r = Replace(r, "30+", "$3")
            r = Replace(r, "40+", "$4")
            r = Replace(r, "50+", "$5")
            r = Replace(r, "60+", "$6")
            r = Replace(r, "70+", "$7")
            r = Replace(r, "80+", "$8")
            r = Replace(r, "90+", "$9")
            r = Replace(r, "00+", "$0")

            r = Replace(r, "0+", "")

            r = Replace(r, "$1", "10+")
            r = Replace(r, "$2", "20+")
            r = Replace(r, "$3", "30+")
            r = Replace(r, "$4", "40+")
            r = Replace(r, "$5", "50+")
            r = Replace(r, "$6", "60+")
            r = Replace(r, "$7", "70+")
            r = Replace(r, "$8", "80+")
            r = Replace(r, "$9", "90+")
            r = Replace(r, "$0", "00+")


            If Comp(r, "0") = "a=b" Then
                Exit For
            End If
        Next
endr:
        Derivation = r
    End Function

    Private Function DemloNum(ByVal n As String) As String
        Dim r As String
        r = stringd(CInt(n), "1")
        DemloNum = r
    End Function
    Private Function BaseConv(ByVal n As String) As String
        Dim z, z2 As Int64
        Dim nu, bi, bo As String
        z = InStr(n, "_")
        nu = Mid(n, 1, z - 1)
        z2 = InStr(Mid(n, z + 1), "_")
        bi = Mid(n, z + 1, z2 - 1)
        z = z + z2
        bo = Mid(n, z + 1)
        BaseConv = convertB(nu, bi, bo)
    End Function
    Private Function convertB(ByVal nu As String, ByVal bi As String, ByVal bo As String) As String
        Dim i, j As Int64
        Dim c As Char
        Dim s As String
        For i = Len(nu) To 1 Step -1
            c = Mid(nu, i, 1)
            s = add(s, Mul(pow(bi, j), cnva(c)))
            j = j + 1
        Next
        convertB = convertN(s, bo)
    End Function
    Private Function convertN(ByVal s As String, ByVal bo As String) As String
        Dim r, m, st, s2 As String
        Dim a(1) As Byte
        Dim y, yb, i As Int32
        m = s
        s2 = Logxb(s, bo)
        yb = ipart(Trim(s2))
        ReDim a(Val(yb))
        Do Until m <= 0
            y = ipart(Trim(Logxb(m, bo)))
            r = pow(bo, y)
rept:
            m = subs(m, r)
            a(Val(y)) = a(Val(y)) + 1
            If Comp(m, r) = "a>b" Then
                GoTo rept
            End If
        Loop

        For i = 0 To Val(yb)
            If a(i) = Nothing Then
                st = "0" & st
            Else
                If a(i) > 9 Then
                    st = acnv(a(i)) & st
                Else
                    st = a(i) & st
                End If
            End If
        Next
        convertN = st
    End Function
    Private Function BellNum(ByVal n As String) As String
        Dim s As String
        Dim i As Int64
        For i = 1 To Val(n)
            s = add(s, StirlingN2(n, i))
        Next i
        BellNum = s
    End Function
    Private Function PronicNum(ByVal n As String) As String
        Dim i As Int64
        PronicNum = Mul(n, add(n, "1"))
    End Function
    Private Function BellTriangle(ByVal n As Integer) As String
        Dim s, r As String
        Dim k, j As Int64
        Dim i(1, 1) As String
        ReDim i(n, n)
        i(1, 1) = "1"
        For k = 2 To n
            i(k, 1) = r
            For j = 2 To k
                r = i(k, j)
                i(k, j) = add(i(k, j - 1), i(k - 1, j - 1))
            Next j
        Next k
        For j = 1 To n
            s = s & i(n, j) & "."
        Next
        BellTriangle = Microsoft.VisualBasic.Left(s, Len(s) - 1)
    End Function
    Private Function LosaTriangle(ByVal n As Integer) As String
        Dim k, j, i As Int64
        Dim s, r As String
        Dim a(1, 1) As String
        ReDim a(n, n)
        r = "1"
        For i = 0 To n
            a(i, 0) = "1"
        Next
        r = 1
        For k = 1 To n
            For j = 1 To k
                If iseven(k) = "1" And isodd(j) = "1" Then
                    r = BinomCoef(subs(div(k, 2, 10), 1), div(subs(j, 1), 2, 10))
                    a(k, j) = subs(add(a(k - 1, j - 1), a(k - 1, j)), r)
                Else
                    a(k, j) = add(a(k - 1, j - 1), a(k - 1, j))
                End If
            Next j
        Next k
        For j = 0 To n
            s = s & a(n, j) & "."
        Next
        LosaTriangle = Microsoft.VisualBasic.Left(s, Len(s) - 1)
    End Function
    Private Function LeibHarmTriangle(ByVal n As Integer) As String
        Dim s, r As String
        Dim k As Int64
        For k = 1 To n
            r = Mul(k, BinomCoef(n, k))
            s = s & "1/" & r & "."
        Next k
        LeibHarmTriangle = Microsoft.VisualBasic.Left(s, Len(s) - 1)
    End Function
    Private Function BernTriangle(ByVal n As Integer) As String
        Dim k As Int64
        Dim s, r As String
        Dim i As Int64
        For k = 0 To n
            r = 0
            For i = 0 To k
                r = add(r, BinomCoef(n, i))
            Next i
            s = s & r & "."
        Next k
        BernTriangle = Microsoft.VisualBasic.Left(s, Len(s) - 1)
    End Function

    Private Function StirlingN2(ByVal n As String, ByVal k As String) As String
        Dim s, s2, s3 As String
        Dim i As Int64
        s = div(1, Factorial(k), 10)
        s3 = 0
        For i = 0 To Val(k)
            s2 = pow(-1, i)
            s2 = Mul(s2, BinomCoef(k, Trim(Str(i))))
            s2 = Mul(s2, pow(subs(k, Trim(Str(i))), n))
            s3 = add(s3, s2)
        Next
        StirlingN2 = RCINT(Mul(s, s3))
    End Function

    Private Function ClarkTriangle(ByVal n As Integer, ByVal f As Integer) As String
        Dim s As String
        Dim k As Integer
        If n = 0 Then
            ClarkTriangle = "0"
        Else
            'If n = 1 Then
            'ClarkTriangle = "6.1"
            'Else
            For k = 0 To n
                s = s & (add(Mul(f, BinomCoef(n, k + 1)), BinomCoef(n - 1, k - 1))) & "."
            Next k
            ClarkTriangle = Microsoft.VisualBasic.Left(s, Len(s) - 1)
            'End If
        End If
    End Function

    Private Function ExpSum(ByVal n As Integer, ByVal x As Integer) As Double
        Dim s As String
        Dim k As Int64
        For k = 0 To n
            s = s + (x ^ k) / Factorial(k)
        Next k
        ExpSum = s
    End Function

    Private Function PochHamSymb(ByVal n As Integer, ByVal x As Integer) As Double
        Dim s As String
        Dim k As Int64
        s = "1"
        For k = 1 To n
            s = Mul(s, Trim(Str((x + k - 1))))
        Next k
        PochHamSymb = s
    End Function
    Private Function FallFactorial(ByVal n As Integer, ByVal x As Integer) As Double
        Dim s As String
        Dim k As Int64
        s = "1"
        For k = n To 1 Step -1
            s = Mul(s, Trim(Str((x - k + 1))))
        Next k
        FallFactorial = s
    End Function
    Private Function TriangNum(ByVal n As Integer) As Double
        TriangNum = (1 / 2) * n * (n + 1)
    End Function
    Private Function SquareNum(ByVal n As Integer) As Double
        SquareNum = n ^ 2
    End Function
    Private Function PentaNum(ByVal n As Integer) As Double
        PentaNum = n * (3 * n - 1) / 2
    End Function
    Private Function HexNum(ByVal n As Integer) As Double
        HexNum = n * (2 * n - 1)
    End Function
    Private Function CenHexNum(ByVal n As Integer) As Double
        CenHexNum = 3 * n ^ 2 + 3 * n + 1
    End Function
    Private Function CubicNum(ByVal n As Integer) As Double
        CubicNum = n ^ 3
    End Function
    Private Function CenCubicNum(ByVal n As Integer) As Double
        CenCubicNum = n ^ 3 + (n - 1) ^ 3
    End Function
    Private Function iseven(ByVal n As String) As Char
        Dim r As String
        r = remn(n, 2)
        If r = "0" Or r = "-0" Then
            iseven = "1"
        Else
            iseven = "0"
        End If
    End Function
    Private Function stringd(ByVal n As Int64, ByVal s As String) As String
        Dim i As Int64
        Dim s2 As String
        For i = 1 To n
            s2 = s2 & s
        Next
        stringd = s2
    End Function
    Private Function isodd(ByVal n As String) As Char
        Dim r As String
        r = remn(n, 2)
        If r = 0 Then
            isodd = "0"
        Else
            isodd = "1"
        End If
    End Function
    Private Function Factors(ByVal n As String) As String
        Dim s, r As String
        Dim j As Int64
        For j = 1 To CInt(n)
            r = remn(n, j)
            If r = "0" Or r = "-0" Then
                s = s & j & ";"
            End If
        Next j
        Factors = s
    End Function
    Private Function SimpContFrac(ByVal n As Integer, ByVal x As Double) As Double
        Dim a, q, z
        Dim s As Int64
        ReDim a(n)
        ReDim q(n)
        ReDim z(n)
        q(0) = x
        a(0) = Floor(q(0))
        If n = 0 Then
            SimpContFrac = a(0)
        Else
            Dim k As Integer
            For k = 1 To n
                q(k) = 1 / (q(k - 1) - a(k - 1))
                a(k) = Floor(1 / (q(k - 1) - a(k - 1)))
                s = s * (x - k + 1)
            Next k
            z(n) = 1 / a(n)
            For k = n - 1 To 1 Step -1
                z(k) = 1 / (a(k) + z(k + 1))
            Next k
            SimpContFrac = z(1) + a(0)
        End If
    End Function
    Private Function GCD2(ByVal a As String, ByVal b As String) As String
        Dim m, s, r As String
        Dim i As Int64
        If Comp(a, b) = "a>b" Then
            m = b
        Else
            m = a
        End If
        s = "1"
        For i = CInt(m) To 1 Step -1
            r = remn(a, Trim(Str(i)))
            If r = "0" Then
                r = remn(b, Trim(Str(i)))
                If r = "0" Then
                    s = Trim(Str(i))
                    Exit For
                End If
            End If
        Next i

        GCD2 = s
    End Function
    Private Function coprime(ByVal m As String, ByVal n As String) As Char
        Dim f As Boolean
        Dim w, r As String
        Dim i As Int64
        If Comp(m, n) = "a>b" Then
            w = m
        Else
            w = n
        End If
        f = True
        For i = 2 To CInt(w)
            r = remn(m, i)
            If r = "0" Or r = "-0" Then
                r = remn(n, i)
                If r = "0" Or r = "-0" Then
                    f = False
                    GoTo en
                End If
            End If
        Next i
en:
        If f = True Then
            coprime = "1"
        Else
            coprime = "0"
        End If
    End Function

    Private Function PowComplexZ(ByVal x As Double, ByVal y As Double, ByVal n As Integer) As String
        Dim k As Int64
        Dim i As Short
        Dim s1, s2, s As String
        s1 = "0"
        For k = 0 To n Step 2
            s1 = s1 + Mul(i, Mul(BinomCoef(n, k), Mul(pow(x, (n - k)), pow(y, k))))
            i = -i
        Next k
        s2 = "0"
        i = 1
        For k = 1 To n Step 2
            s2 = s2 + Mul(i, Mul(BinomCoef(n, k), Mul(pow(x, (n - k)), pow(y, k))))
            i = -i
        Next k
        If s2 < 0 Then
            s = s1 & " - i(" & -s2 & ")"
        Else
            s = s1 & " + i(" & s2 & ")"
        End If
        PowComplexZ = s
    End Function
    Public sw, sw2 As Byte
    Private Sub Text1_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text1.MouseHover
        If sw = 1 Then
            Label2.Visible = True
        End If
    End Sub

    Private Sub Text1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text1.MouseLeave
        sw = 1
        Label2.Visible = False
    End Sub

    Private Sub Text1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text1.Click
        sw = 0
        Label2.Visible = False
    End Sub

    Private Sub Text1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Text1.KeyPress
        sw = 0
        Label2.Visible = False
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' MsgBox(Liouville(4))
        'MsgBox(Rut(9, 2, 10))
        'MsgBox(add("96", "4"))
        sw = 1
        sw2 = 1
        Label2.Visible = False
        Label3.Visible = False
        'MsgBox(repunit(33, 10))
    End Sub
    Private Sub text2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles text2.KeyPress
        sw2 = 0
        Label3.Visible = False
    End Sub

    Private Sub text2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles text2.Click
        sw2 = 0
        Label3.Visible = False
    End Sub
    Private Sub text2_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles text2.MouseHover
        If sw2 = 1 Then
            Label3.Visible = True
        End If
    End Sub

    Private Sub text2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles text2.MouseLeave
        sw2 = 1
        Label3.Visible = False
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub Text1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Text1.MouseDown
        Dim pos As New Point
        pos.X = e.X
        pos.Y = e.Y
        If e.Button = MouseButtons.Right Then
            ContextMenu1.Show(Text1, pos)
        End If
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        Text1.SelectedText = ChrW(12313) & " " & ChrW(12314)
    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Text1.SelectedText = ChrW(12315) & " " & ChrW(12316)
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Text1.SelectedText = "#CatTriangle()"
    End Sub

    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click
        Text1.SelectedText = "#BernTriangle()"
    End Sub

    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click
        Text1.SelectedText = "#GCD(;)"
    End Sub

    Private Sub MenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem8.Click
        Text1.SelectedText = "#Coprime(;)"
    End Sub

    Private Sub MenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem9.Click
        Text1.SelectedText = "#BinomCoef(;)"
    End Sub

    Private Sub MenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem10.Click
        Text1.SelectedText = "#CentBiCoef()"
    End Sub

    Private Sub MenuItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem11.Click
        Text1.SelectedText = "#PasTriangle()"
    End Sub

    Private Sub MenuItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem12.Click
        Text1.SelectedText = "#CenCubicNum()"
    End Sub

    Private Sub MenuItem13_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem13.Click
        Text1.SelectedText = "#CenHexNum()"
    End Sub

    Private Sub MenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem14.Click
        Text1.SelectedText = "#ClarkTriangle(;)"
    End Sub

    Private Sub MenuItem15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem15.Click
        Text1.SelectedText = "#StirlingN2(;)"
    End Sub

    Private Sub MenuItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem16.Click
        Text1.SelectedText = "#BellNum()"
    End Sub

    Private Sub MenuItem17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem17.Click
        Text1.SelectedText = "[ " & "]"
    End Sub

    Private Sub MenuItem18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem18.Click
        Text1.SelectedText = "| " & "|"
    End Sub

    Private Sub MenuItem19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem19.Click
        Text1.SelectedText = "!"
    End Sub

    Private Sub MenuItem20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem20.Click
        Text1.SelectedText = "!!"
    End Sub

    Private Sub MenuItem21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem21.Click
        Text1.SelectedText = "#BellTriangle()"
    End Sub

    Private Sub MenuItem22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem22.Click
        Text1.SelectedText = "#LosaTriangle()"
    End Sub

    Private Sub Command1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub MenuItem23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem23.Click
        Text1.SelectedText = "#LeibHarmTriangle()"
    End Sub

    Private Sub MenuItem24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem24.Click
        Text1.SelectedText = "#PronicNum()"
    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click

    End Sub

    Private Sub MenuItem27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem27.Click
        Text1.SelectedText = "#IsEven()"
    End Sub

    Private Sub MenuItem28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem28.Click
        Text1.SelectedText = "#IsOdd()"
    End Sub

    Private Sub MenuItem30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem30.Click
        Text1.SelectedText = "#DemloNum()"
    End Sub

    Private Sub MenuItem31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem31.Click
        Text1.SelectedText = "#Repunit(;)"
    End Sub

    Private Sub MenuItem33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem33.Click
        Text1.SelectedText = "#ToRoman()"
    End Sub


    Private Sub MenuItem35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem35.Click
        Text1.SelectedText = "Ī"
    End Sub

    Private Sub MenuItem36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem36.Click
        Text1.SelectedText = ChrW(12289)
    End Sub

    Private Sub MenuItem37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem37.Click
        Text1.SelectedText = ChrW(12288)
    End Sub

    Private Sub MenuItem38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem38.Click
        Text1.SelectedText = ChrW(12300)
    End Sub

    Private Sub MenuItem39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem39.Click
        Text1.SelectedText = "||"
    End Sub

    Private Sub MenuItem40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem40.Click
        Text1.SelectedText = "#FromRoman()"
    End Sub

    Private Sub MenuItem41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub MenuItem43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem43.Click
        Text1.SelectedText = "{F0001_Number_InputBase_OutputBase}"
    End Sub

    Private Sub MenuItem41_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem41.Click
        Text1.SelectedText = "#IsPrime()"
    End Sub

    Private Sub MenuItem44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem44.Click
        Text1.SelectedText = "#Perrin(;0)"
    End Sub

    Private Sub MenuItem47_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem47.Click
        Text1.SelectAll()
    End Sub

    Private Sub MenuItem45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem45.Click
        Text1.Copy()
    End Sub

    Private Sub MenuItem46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem46.Click
        Text1.Paste()
    End Sub

    Private Sub MenuItem48_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem48.Click
        Text1.Undo()
    End Sub

    Private Sub MenuItem49_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem49.Click
        Text1.SelectedText = "#Fibonacci(;0)"
    End Sub

    Private Sub MenuItem53_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem53.Click
        Text1.SelectedText = "e"
    End Sub

    Private Sub MenuItem51_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem51.Click
        Text1.SelectedText = ChrW(216)
    End Sub

    Private Sub MenuItem54_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Text1.SelectedText = "{F0002_Expression}"
    End Sub

    Private Sub MenuItem52_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem52.Click
        Text1.SelectedText = ChrW(928)
    End Sub

    Private Sub MenuItem55_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem55.Click
        Text1.SelectedText = "#EulerN(;)"
    End Sub

    Private Sub MenuItem56_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem56.Click
        Text1.SelectedText = "#IsInteger()"
    End Sub

    Private Sub MenuItem59_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem59.Click
        Text1.SelectedText = "If <condition> THEN <expr> ELSE <expr2> END"
    End Sub

    Private Sub MenuItem58_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem58.Click
        Text1.SelectedText = "If <condition> THEN <expr> END"
    End Sub

    Private Sub MenuItem60_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem60.Click
        Text1.SelectedText = ChrW(947)
    End Sub

    Private Sub MenuItem61_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem61.Click
        Text1.SelectedText = ChrW(12302)
    End Sub

    Private Sub MenuItem62_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem62.Click
        Text1.SelectedText = ChrW(12303)
    End Sub

    Private Sub MenuItem63_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem63.Click
        Text1.SelectedText = ChrW(12304)
    End Sub

    Private Sub MenuItem64_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem64.Click
        Text1.SelectedText = ChrW(12305)
    End Sub

    Private Sub MenuItem65_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem65.Click
        Text1.SelectedText = ChrW(12306)
    End Sub

    Private Sub MenuItem66_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem66.Click
        Text1.SelectedText = ChrW(12307)
    End Sub

    Private Sub MenuItem67_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem67.Click
        Text1.SelectedText = ChrW(12308)
    End Sub

    Private Sub MenuItem68_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem68.Click
        Text1.SelectedText = ChrW(12309)
    End Sub

    Private Sub MenuItem69_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem69.Click
        Text1.SelectedText = ChrW(12310)
    End Sub

    Private Sub MenuItem70_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem70.Click
        Text1.SelectedText = ChrW(12311)
    End Sub

    Private Sub MenuItem71_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem71.Click
        Text1.SelectedText = ChrW(12312)
    End Sub

    Private Sub MenuItem72_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem72.Click
        Text1.SelectedText = "#Liouville()"
    End Sub

    Private Sub MenuItem73_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem73.Click
        Text1.SelectedText = "#Parity()"
    End Sub

    Private Sub MenuItem74_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem74.Click
        Text1.SelectedText = "#ThueMorse()"
    End Sub

    Private Sub Command2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Command2.MouseDown
        PictureBox1.Visible = True
        PictureBox2.Visible = True
        PictureBox3.Visible = True
        text2.Visible = False
    End Sub

    Private Sub MenuItem77_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem77.Click
        Text1.SelectedText = "#cos(;d)"
    End Sub

    Private Sub MenuItem76_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem76.Click
        Text1.SelectedText = "#sin(;d)"
    End Sub

    Private Sub MenuItem78_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem78.Click
        Text1.SelectedText = "#tan(;d)"
    End Sub

    Private Sub MenuItem79_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem79.Click
        Text1.SelectedText = "#ctg(;d)"
    End Sub

    Private Sub MenuItem80_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem80.Click
        Text1.SelectedText = ChrW(8730)
    End Sub

    Private Sub MenuItem81_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem81.Click
        Text1.SelectedText = "#sec(;d)"
    End Sub

    Private Sub MenuItem82_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem82.Click
        Text1.SelectedText = "#csc(;d)"
    End Sub

    Private Sub MenuItem83_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem83.Click
        Text1.SelectedText = "#sinh(;d)"
    End Sub

    Private Sub MenuItem84_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem84.Click
        Text1.SelectedText = "#cosh(;d)"
    End Sub

    Private Sub MenuItem85_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem85.Click
        Text1.SelectedText = "#sech(;d)"
    End Sub

    Private Sub MenuItem86_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem86.Click
        Text1.SelectedText = "#csch(;d)"
    End Sub

    Private Sub MenuItem54_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem54.Click
        Text1.SelectedText = "{F0002_Expression}"
    End Sub

    Private Sub MenuItem88_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem88.Click
        Text1.SelectedText = "#Max(n1:...:nN)"
    End Sub

    Private Sub MenuItem89_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem89.Click
        Text1.SelectedText = "#Min(n1:...:nN)"
    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click
        Dim a As New Shell32.Shell
        a.Open("http://ezer0.net")
        'Microsoft.VisualBasic.Shell("explorer.exe http://ezer0.net", AppWinStyle.MaximizedFocus, False, )
    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click

    End Sub

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label9.Click
        Dim a As New Shell32.Shell
        a.Open("mailto:evg_zel@mail.ru")
        ' a.FileRun()
    End Sub

    Private Sub Label11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.Click
        Label11.Text = "BACK"
        Panel1.Visible = False
    End Sub

    Private Sub PictureBox7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox7.Click
        Randomize()
        Dim a(2) As Image
        a(0) = PictureBox8.Image
        a(1) = PictureBox9.Image
        a(2) = PictureBox10.Image
        Dim r As Int16
        r = Int(Rnd() * 3)
        PictureBox7.Image = a(r)
    End Sub

    Private Sub PictureBox6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox6.Click
        Randomize()
        Dim a(2) As Image
        a(0) = PictureBox8.Image
        a(1) = PictureBox9.Image
        a(2) = PictureBox10.Image
        Dim r As Int16
        r = Int(Rnd() * 3)
        PictureBox6.Image = a(r)
    End Sub

    Private Sub PictureBox5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox5.Click
        Randomize()
        Dim a(2) As Image
        a(0) = PictureBox8.Image
        a(1) = PictureBox9.Image
        a(2) = PictureBox10.Image
        Dim r As Int16
        r = Int(Rnd() * 3)
        PictureBox5.Image = a(r)
    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

    End Sub

    Private Sub Label1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Panel1.Visible = True
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click
        Dim a As New Shell32.Shell
        Dim b As String
        'MsgBox()
        b = Microsoft.VisualBasic.CurDir() & "\MATEZ.mht"
        a.Open(b)
    End Sub

    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click
        MsgBox(ArabicToMayan("20"))
    End Sub

    Private Sub MenuItem91_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem91.Click
        Text1.SelectedText = ChrW(&H301D)
    End Sub

    Private Sub MenuItem92_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem92.Click
        Text1.SelectedText = ChrW(&H301E)
    End Sub

    Private Sub MenuItem93_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem93.Click
        Text1.SelectedText = ChrW(&H301F)
    End Sub

    Private Sub MenuItem94_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem94.Click
        Text1.SelectedText = ChrW(&H3020)
    End Sub

    Private Sub MenuItem95_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem95.Click
        Text1.SelectedText = ChrW(&H3021)
    End Sub

    Private Sub MenuItem96_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem96.Click
        Text1.SelectedText = ChrW(&H3022)
    End Sub

    Private Sub MenuItem97_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem97.Click
        Text1.SelectedText = ChrW(&H3023)
    End Sub

    Private Sub MenuItem98_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem98.Click
        Text1.SelectedText = ChrW(&H3024)
    End Sub

    Private Sub MenuItem99_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem99.Click
        Text1.SelectedText = ChrW(&H3025)
    End Sub

    Private Sub MenuItem100_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem100.Click
        Text1.SelectedText = ChrW(&H3026)
    End Sub

    Private Sub MenuItem101_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem101.Click
        Text1.SelectedText = ChrW(&H3027)
    End Sub

    Private Sub MenuItem102_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem102.Click
        Text1.SelectedText = ChrW(&H3028)
    End Sub

    Private Sub MenuItem103_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem103.Click
        Text1.SelectedText = ChrW(&H3029)
    End Sub

    Private Sub MenuItem104_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem104.Click
        Text1.SelectedText = ChrW(&H3030)
    End Sub

    Private Sub MenuItem105_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem105.Click
        Text1.SelectedText = ChrW(&H3031)
    End Sub

    Private Sub MenuItem106_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem106.Click
        Text1.SelectedText = ChrW(&H3032)
    End Sub

    Private Sub MenuItem107_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem107.Click
        Text1.SelectedText = ChrW(&H3033)
    End Sub

    Private Sub MenuItem108_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem108.Click
        Text1.SelectedText = ChrW(&H3034)
    End Sub

    Private Sub MenuItem109_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem109.Click
        Text1.SelectedText = ChrW(&H3035)
    End Sub

    Private Sub MenuItem110_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem110.Click
        Text1.SelectedText = ChrW(&H3036)
    End Sub

    Private Sub MenuItem111_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem111.Click
        Text1.SelectedText = "#ToMayan()"
    End Sub

    Private Sub MenuItem34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem34.Click

    End Sub

    Private Sub MenuItem112_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem112.Click
        Text1.SelectedText = "#NumText()"
    End Sub

    Private Sub MenuItem113_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem113.Click
        Text1.SelectedText = "#ToMiken()"
    End Sub

    Private Sub MenuItem115_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem115.Click
        Text1.SelectedText = ChrW(&H3037)
    End Sub

    Private Sub MenuItem116_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem116.Click
        Text1.SelectedText = ChrW(&H3038)
    End Sub

    Private Sub MenuItem117_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem117.Click
        Text1.SelectedText = ChrW(&H3039)
    End Sub

    Private Sub MenuItem118_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem118.Click
        Text1.SelectedText = ChrW(&H303A)
    End Sub

    Private Sub MenuItem119_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem119.Click
        Text1.SelectedText = ChrW(&H303C)
    End Sub

    Private Sub MenuItem120_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem120.Click
        Text1.SelectedText = "#FromMiken()"
    End Sub

    Private Sub MenuItem121_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem121.Click
        Text1.SelectedText = "#ToCyril()"
    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click

    End Sub

    Private Sub MenuItem128_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem128.Click
        Text1.SelectedText = "#tanh(;d)"
    End Sub

    Private Sub MenuItem129_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem129.Click
        Text1.SelectedText = "#ctgh(;d)"
    End Sub

    Private Sub MenuItem122_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem122.Click
        Text1.SelectedText = "#asin()"
    End Sub

    Private Sub MenuItem123_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem123.Click
        Text1.SelectedText = "#acos()"
    End Sub

    Private Sub MenuItem124_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem124.Click
        Text1.SelectedText = "#atg()"
    End Sub

    Private Sub MenuItem125_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem125.Click
        Text1.SelectedText = "#actg()"
    End Sub

    Private Sub MenuItem126_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem126.Click
        Text1.SelectedText = "#asec()"
    End Sub

    Private Sub MenuItem127_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem127.Click
        Text1.SelectedText = "#acsc()"
    End Sub

    Private Sub MenuItem130_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem130.Click
        Text1.SelectedText = ChrW(&H25)
    End Sub

    Private Sub MenuItem131_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem131.Click
        Text1.SelectedText = "#DigiRoot()"
    End Sub

    Private Sub MenuItem133_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem133.Click
        Text1.SelectedText = "#DigiSum(;)"
    End Sub

    Private Sub MenuItem134_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem134.Click
        Text1.SelectedText = "#DigiBlock(;;)"
    End Sub

    Private Sub MenuItem135_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem135.Click
        Text1.SelectedText = ChrW(&H3B2)
    End Sub
    Private Sub MenuItem136_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem136.Click
        Text1.SelectedText = "#Round(;)"
    End Sub
    Private Sub MenuItem138_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem138.Click
        Text1.SelectedText = "#Div(;;10)"
    End Sub
    Private Sub MenuItem140_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem140.Click
        Text1.SelectedText = "#Root(;;10)"
    End Sub
    Private Sub MenuItem141_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem141.Click
        Text1.SelectedText = "#Log(;;10)"
    End Sub

    Private Sub MenuItem142_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem142.Click
        Text1.SelectedText = "#Ln(10)"
    End Sub

    Private Sub MenuItem143_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem143.Click
        Text1.SelectedText = "#Log(;2;10)"
    End Sub

    Private Sub MenuItem144_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem144.Click
        Text1.SelectedText = "#Log(;10;10)"
    End Sub

    Private Sub MenuItem145_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem145.Click
        Text1.SelectedText = "#Ln2(10)"
    End Sub

    Private Sub MenuItem146_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem146.Click
        Text1.SelectedText = "#Ln2c()"
    End Sub


    Private Sub PictureBox1_EnabledChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.EnabledChanged

    End Sub

    Private Sub PictureBox1_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.VisibleChanged
    End Sub

    Private Sub MenuItem147_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem147.Click
        Text1.SelectedText = "#Deriv(;1)"
    End Sub

    Private Sub ContextMenu1_Popup(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContextMenu1.Popup

    End Sub

    Private Sub MenuItem148_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem148.Click
        Text1.SelectedText = "#Var()"
    End Sub

    Private Sub MenuItem149_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem149.Click
        MenuItem149.Checked = Not MenuItem149.Checked
    End Sub

    Private Sub Text1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Text1.TextChanged

    End Sub

    Private Sub MenuItem150_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem150.Click
        Text1.SelectedText = "#Pow(;;10)"
    End Sub
End Class