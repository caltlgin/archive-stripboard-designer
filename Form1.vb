Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents ListBox3 As System.Windows.Forms.ListBox
    Friend WithEvents ListBox4 As System.Windows.Forms.ListBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents listbox5 As System.Windows.Forms.ListBox
    Friend WithEvents listbox7 As System.Windows.Forms.ListBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents listbox8 As System.Windows.Forms.ListBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents listbox13 As System.Windows.Forms.ListBox
    Friend WithEvents listbox14 As System.Windows.Forms.ListBox
    Friend WithEvents listbox15 As System.Windows.Forms.ListBox
    Friend WithEvents listbox16 As System.Windows.Forms.ListBox
    Friend WithEvents listbox17 As System.Windows.Forms.ListBox
    Friend WithEvents listbox18 As System.Windows.Forms.ListBox
    Friend WithEvents ListBox21 As System.Windows.Forms.ListBox
    Friend WithEvents listbox22 As System.Windows.Forms.ListBox
    Friend WithEvents ListBox24 As System.Windows.Forms.ListBox
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents ComboBox7 As System.Windows.Forms.ComboBox
    Friend WithEvents ListBox11 As System.Windows.Forms.ListBox
    Friend WithEvents ListBox12 As System.Windows.Forms.ListBox
    Friend WithEvents ListBox10 As System.Windows.Forms.ListBox
    Friend WithEvents ListBox9 As System.Windows.Forms.ListBox
    Friend WithEvents ListBox26 As System.Windows.Forms.ListBox
    Friend WithEvents TextBox26 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox27 As System.Windows.Forms.TextBox
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents TextBox35 As System.Windows.Forms.TextBox
    Friend WithEvents Button40 As System.Windows.Forms.Button
    Friend WithEvents TextBox14 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox15 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox20 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox21 As System.Windows.Forms.TextBox
    Friend WithEvents ListBox28 As System.Windows.Forms.ListBox
    Friend WithEvents ListBox23 As System.Windows.Forms.ListBox
    Friend WithEvents ListBox25 As System.Windows.Forms.ListBox
    Friend WithEvents ListBox27 As System.Windows.Forms.ListBox
    Friend WithEvents Button42 As System.Windows.Forms.Button
    Friend WithEvents Button41 As System.Windows.Forms.Button
    Friend WithEvents Button43 As System.Windows.Forms.Button
    Friend WithEvents Button27 As System.Windows.Forms.Button
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox8 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox9 As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Button15 As System.Windows.Forms.Button
    Friend WithEvents TextBox24 As System.Windows.Forms.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox22 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox18 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox23 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox32 As System.Windows.Forms.TextBox
    Friend WithEvents ListBox20 As System.Windows.Forms.ListBox
    Friend WithEvents ListBox19 As System.Windows.Forms.ListBox
    Friend WithEvents TextBox19 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox33 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox10 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox17 As System.Windows.Forms.TextBox
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button9 As System.Windows.Forms.Button
    Friend WithEvents ListBox2 As System.Windows.Forms.ListBox
    'Friend WithEvents ListBox5 As System.Windows.Forms.ListBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.ListBox3 = New System.Windows.Forms.ListBox
        Me.ListBox4 = New System.Windows.Forms.ListBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.TextBox5 = New System.Windows.Forms.TextBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.ListBox21 = New System.Windows.Forms.ListBox
        Me.ListBox24 = New System.Windows.Forms.ListBox
        Me.listbox5 = New System.Windows.Forms.ListBox
        Me.Button8 = New System.Windows.Forms.Button
        Me.ComboBox7 = New System.Windows.Forms.ComboBox
        Me.ListBox11 = New System.Windows.Forms.ListBox
        Me.ListBox12 = New System.Windows.Forms.ListBox
        Me.ListBox10 = New System.Windows.Forms.ListBox
        Me.ListBox9 = New System.Windows.Forms.ListBox
        Me.ListBox26 = New System.Windows.Forms.ListBox
        Me.TextBox26 = New System.Windows.Forms.TextBox
        Me.TextBox27 = New System.Windows.Forms.TextBox
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.TextBox35 = New System.Windows.Forms.TextBox
        Me.Button40 = New System.Windows.Forms.Button
        Me.TextBox14 = New System.Windows.Forms.TextBox
        Me.TextBox15 = New System.Windows.Forms.TextBox
        Me.TextBox20 = New System.Windows.Forms.TextBox
        Me.TextBox21 = New System.Windows.Forms.TextBox
        Me.ListBox28 = New System.Windows.Forms.ListBox
        Me.ListBox23 = New System.Windows.Forms.ListBox
        Me.ListBox25 = New System.Windows.Forms.ListBox
        Me.ListBox27 = New System.Windows.Forms.ListBox
        Me.Button42 = New System.Windows.Forms.Button
        Me.Button41 = New System.Windows.Forms.Button
        Me.Button43 = New System.Windows.Forms.Button
        Me.Button27 = New System.Windows.Forms.Button
        Me.TextBox7 = New System.Windows.Forms.TextBox
        Me.TextBox8 = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.TextBox9 = New System.Windows.Forms.TextBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button7 = New System.Windows.Forms.Button
        Me.Button15 = New System.Windows.Forms.Button
        Me.TextBox24 = New System.Windows.Forms.TextBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Button4 = New System.Windows.Forms.Button
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.GroupBox7 = New System.Windows.Forms.GroupBox
        Me.Button9 = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.TextBox22 = New System.Windows.Forms.TextBox
        Me.TextBox18 = New System.Windows.Forms.TextBox
        Me.TextBox23 = New System.Windows.Forms.TextBox
        Me.TextBox32 = New System.Windows.Forms.TextBox
        Me.ListBox20 = New System.Windows.Forms.ListBox
        Me.ListBox19 = New System.Windows.Forms.ListBox
        Me.ListBox2 = New System.Windows.Forms.ListBox
        Me.TextBox19 = New System.Windows.Forms.TextBox
        Me.TextBox33 = New System.Windows.Forms.TextBox
        Me.TextBox10 = New System.Windows.Forms.TextBox
        Me.TextBox17 = New System.Windows.Forms.TextBox
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(6, 1072)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(104, 22)
        Me.TextBox1.TabIndex = 5
        Me.TextBox1.Text = "TextBox1"
        Me.TextBox1.Visible = False
        '
        'ListBox3
        '
        Me.ListBox3.BackColor = System.Drawing.Color.AntiqueWhite
        Me.ListBox3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBox3.ItemHeight = 16
        Me.ListBox3.Location = New System.Drawing.Point(1246, 129)
        Me.ListBox3.Name = "ListBox3"
        Me.ListBox3.Size = New System.Drawing.Size(275, 372)
        Me.ListBox3.TabIndex = 8
        '
        'ListBox4
        '
        Me.ListBox4.BackColor = System.Drawing.Color.AntiqueWhite
        Me.ListBox4.HorizontalScrollbar = True
        Me.ListBox4.ItemHeight = 16
        Me.ListBox4.Location = New System.Drawing.Point(1255, 149)
        Me.ListBox4.Name = "ListBox4"
        Me.ListBox4.Size = New System.Drawing.Size(193, 212)
        Me.ListBox4.Sorted = True
        Me.ListBox4.TabIndex = 10
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(6, 1104)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(104, 22)
        Me.TextBox2.TabIndex = 11
        Me.TextBox2.Text = "TextBox2"
        Me.TextBox2.Visible = False
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(10, 1033)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(100, 22)
        Me.TextBox3.TabIndex = 12
        Me.TextBox3.Text = "TextBox3"
        Me.TextBox3.Visible = False
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(10, 966)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(100, 22)
        Me.TextBox4.TabIndex = 13
        Me.TextBox4.Text = "TextBox4"
        Me.TextBox4.Visible = False
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(10, 990)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(100, 22)
        Me.TextBox5.TabIndex = 19
        Me.TextBox5.Text = "TextBox5"
        Me.TextBox5.Visible = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.Ivory
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PictureBox1.Location = New System.Drawing.Point(12, 131)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(1175, 823)
        Me.PictureBox1.TabIndex = 40
        Me.PictureBox1.TabStop = False
        '
        'ListBox21
        '
        Me.ListBox21.BackColor = System.Drawing.Color.AntiqueWhite
        Me.ListBox21.ItemHeight = 16
        Me.ListBox21.Location = New System.Drawing.Point(1658, 1033)
        Me.ListBox21.Name = "ListBox21"
        Me.ListBox21.Size = New System.Drawing.Size(43, 20)
        Me.ListBox21.TabIndex = 86
        Me.ListBox21.Visible = False
        '
        'ListBox24
        '
        Me.ListBox24.BackColor = System.Drawing.Color.AntiqueWhite
        Me.ListBox24.ItemHeight = 16
        Me.ListBox24.Location = New System.Drawing.Point(1673, 1038)
        Me.ListBox24.Name = "ListBox24"
        Me.ListBox24.Size = New System.Drawing.Size(187, 100)
        Me.ListBox24.TabIndex = 93
        '
        'listbox5
        '
        Me.listbox5.BackColor = System.Drawing.Color.AntiqueWhite
        Me.listbox5.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.listbox5.ItemHeight = 16
        Me.listbox5.Location = New System.Drawing.Point(1637, 1012)
        Me.listbox5.Name = "listbox5"
        Me.listbox5.Size = New System.Drawing.Size(293, 148)
        Me.listbox5.TabIndex = 95
        '
        'Button8
        '
        Me.Button8.BackColor = System.Drawing.Color.Snow
        Me.Button8.Location = New System.Drawing.Point(14, 59)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(55, 27)
        Me.Button8.TabIndex = 99
        Me.Button8.Text = "Clear"
        Me.Button8.UseVisualStyleBackColor = False
        '
        'ComboBox7
        '
        Me.ComboBox7.Location = New System.Drawing.Point(6, 18)
        Me.ComboBox7.Name = "ComboBox7"
        Me.ComboBox7.Size = New System.Drawing.Size(260, 24)
        Me.ComboBox7.TabIndex = 103
        Me.ComboBox7.Text = "ComboBox7"
        '
        'ListBox11
        '
        Me.ListBox11.BackColor = System.Drawing.Color.AntiqueWhite
        Me.ListBox11.ItemHeight = 16
        Me.ListBox11.Location = New System.Drawing.Point(1650, 1088)
        Me.ListBox11.Name = "ListBox11"
        Me.ListBox11.Size = New System.Drawing.Size(47, 20)
        Me.ListBox11.TabIndex = 58
        Me.ListBox11.Visible = False
        '
        'ListBox12
        '
        Me.ListBox12.BackColor = System.Drawing.Color.AntiqueWhite
        Me.ListBox12.ItemHeight = 16
        Me.ListBox12.Location = New System.Drawing.Point(1703, 1114)
        Me.ListBox12.Name = "ListBox12"
        Me.ListBox12.Size = New System.Drawing.Size(72, 20)
        Me.ListBox12.TabIndex = 59
        Me.ListBox12.Visible = False
        '
        'ListBox10
        '
        Me.ListBox10.BackColor = System.Drawing.Color.AntiqueWhite
        Me.ListBox10.ItemHeight = 16
        Me.ListBox10.Location = New System.Drawing.Point(1781, 1114)
        Me.ListBox10.Name = "ListBox10"
        Me.ListBox10.Size = New System.Drawing.Size(70, 20)
        Me.ListBox10.TabIndex = 62
        Me.ListBox10.Visible = False
        '
        'ListBox9
        '
        Me.ListBox9.BackColor = System.Drawing.Color.AntiqueWhite
        Me.ListBox9.ItemHeight = 16
        Me.ListBox9.Location = New System.Drawing.Point(1650, 1114)
        Me.ListBox9.Name = "ListBox9"
        Me.ListBox9.Size = New System.Drawing.Size(44, 20)
        Me.ListBox9.TabIndex = 52
        Me.ListBox9.Visible = False
        '
        'ListBox26
        '
        Me.ListBox26.BackColor = System.Drawing.Color.AntiqueWhite
        Me.ListBox26.ItemHeight = 16
        Me.ListBox26.Location = New System.Drawing.Point(1791, 1062)
        Me.ListBox26.Name = "ListBox26"
        Me.ListBox26.Size = New System.Drawing.Size(69, 20)
        Me.ListBox26.Sorted = True
        Me.ListBox26.TabIndex = 125
        Me.ListBox26.Visible = False
        '
        'TextBox26
        '
        Me.TextBox26.Location = New System.Drawing.Point(1124, 46)
        Me.TextBox26.Name = "TextBox26"
        Me.TextBox26.Size = New System.Drawing.Size(56, 22)
        Me.TextBox26.TabIndex = 129
        Me.TextBox26.Text = "1985"
        Me.TextBox26.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.TextBox26.Visible = False
        '
        'TextBox27
        '
        Me.TextBox27.Location = New System.Drawing.Point(1124, 77)
        Me.TextBox27.Name = "TextBox27"
        Me.TextBox27.Size = New System.Drawing.Size(56, 22)
        Me.TextBox27.TabIndex = 128
        Me.TextBox27.Text = "930"
        Me.TextBox27.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.TextBox27.Visible = False
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.Color.AntiqueWhite
        Me.ListBox1.HorizontalScrollbar = True
        Me.ListBox1.ItemHeight = 16
        Me.ListBox1.Location = New System.Drawing.Point(1637, 1140)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(59, 20)
        Me.ListBox1.TabIndex = 134
        Me.ListBox1.Visible = False
        '
        'TextBox35
        '
        Me.TextBox35.Location = New System.Drawing.Point(1103, 17)
        Me.TextBox35.Name = "TextBox35"
        Me.TextBox35.Size = New System.Drawing.Size(77, 22)
        Me.TextBox35.TabIndex = 155
        Me.TextBox35.Text = "2"
        Me.TextBox35.Visible = False
        '
        'Button40
        '
        Me.Button40.BackColor = System.Drawing.Color.Snow
        Me.Button40.Location = New System.Drawing.Point(11, 25)
        Me.Button40.Name = "Button40"
        Me.Button40.Size = New System.Drawing.Size(58, 26)
        Me.Button40.TabIndex = 159
        Me.Button40.Text = "New"
        Me.Button40.UseVisualStyleBackColor = False
        '
        'TextBox14
        '
        Me.TextBox14.Location = New System.Drawing.Point(1029, 45)
        Me.TextBox14.Name = "TextBox14"
        Me.TextBox14.Size = New System.Drawing.Size(96, 22)
        Me.TextBox14.TabIndex = 163
        Me.TextBox14.Visible = False
        '
        'TextBox15
        '
        Me.TextBox15.Location = New System.Drawing.Point(1029, 17)
        Me.TextBox15.Name = "TextBox15"
        Me.TextBox15.Size = New System.Drawing.Size(98, 22)
        Me.TextBox15.TabIndex = 164
        Me.TextBox15.Visible = False
        '
        'TextBox20
        '
        Me.TextBox20.Location = New System.Drawing.Point(1137, 45)
        Me.TextBox20.Name = "TextBox20"
        Me.TextBox20.Size = New System.Drawing.Size(43, 22)
        Me.TextBox20.TabIndex = 166
        Me.TextBox20.Visible = False
        '
        'TextBox21
        '
        Me.TextBox21.Location = New System.Drawing.Point(1029, 74)
        Me.TextBox21.Name = "TextBox21"
        Me.TextBox21.Size = New System.Drawing.Size(43, 22)
        Me.TextBox21.TabIndex = 167
        Me.TextBox21.Visible = False
        '
        'ListBox28
        '
        Me.ListBox28.BackColor = System.Drawing.Color.AntiqueWhite
        Me.ListBox28.ItemHeight = 16
        Me.ListBox28.Location = New System.Drawing.Point(1707, 1062)
        Me.ListBox28.Name = "ListBox28"
        Me.ListBox28.Size = New System.Drawing.Size(78, 20)
        Me.ListBox28.TabIndex = 169
        Me.ListBox28.Visible = False
        '
        'ListBox23
        '
        Me.ListBox23.BackColor = System.Drawing.Color.AntiqueWhite
        Me.ListBox23.ItemHeight = 16
        Me.ListBox23.Location = New System.Drawing.Point(1637, 1062)
        Me.ListBox23.Name = "ListBox23"
        Me.ListBox23.Size = New System.Drawing.Size(64, 20)
        Me.ListBox23.TabIndex = 172
        Me.ListBox23.Visible = False
        '
        'ListBox25
        '
        Me.ListBox25.BackColor = System.Drawing.Color.AntiqueWhite
        Me.ListBox25.ItemHeight = 16
        Me.ListBox25.Location = New System.Drawing.Point(1781, 1088)
        Me.ListBox25.Name = "ListBox25"
        Me.ListBox25.Size = New System.Drawing.Size(90, 20)
        Me.ListBox25.TabIndex = 173
        Me.ListBox25.Visible = False
        '
        'ListBox27
        '
        Me.ListBox27.BackColor = System.Drawing.Color.AntiqueWhite
        Me.ListBox27.ItemHeight = 16
        Me.ListBox27.Location = New System.Drawing.Point(1700, 1088)
        Me.ListBox27.Name = "ListBox27"
        Me.ListBox27.Size = New System.Drawing.Size(75, 20)
        Me.ListBox27.TabIndex = 174
        Me.ListBox27.Visible = False
        '
        'Button42
        '
        Me.Button42.BackColor = System.Drawing.Color.Snow
        Me.Button42.ForeColor = System.Drawing.Color.HotPink
        Me.Button42.Location = New System.Drawing.Point(6, 22)
        Me.Button42.Name = "Button42"
        Me.Button42.Size = New System.Drawing.Size(69, 26)
        Me.Button42.TabIndex = 175
        Me.Button42.Text = "Generic"
        Me.Button42.UseVisualStyleBackColor = False
        '
        'Button41
        '
        Me.Button41.BackColor = System.Drawing.Color.Snow
        Me.Button41.Location = New System.Drawing.Point(6, 21)
        Me.Button41.Name = "Button41"
        Me.Button41.Size = New System.Drawing.Size(106, 25)
        Me.Button41.TabIndex = 176
        Me.Button41.Text = "Full Break"
        Me.Button41.UseVisualStyleBackColor = False
        '
        'Button43
        '
        Me.Button43.BackColor = System.Drawing.Color.Snow
        Me.Button43.Location = New System.Drawing.Point(75, 25)
        Me.Button43.Name = "Button43"
        Me.Button43.Size = New System.Drawing.Size(56, 27)
        Me.Button43.TabIndex = 178
        Me.Button43.Text = "Load"
        Me.Button43.UseVisualStyleBackColor = False
        '
        'Button27
        '
        Me.Button27.BackColor = System.Drawing.Color.Snow
        Me.Button27.Location = New System.Drawing.Point(6, 55)
        Me.Button27.Name = "Button27"
        Me.Button27.Size = New System.Drawing.Size(106, 27)
        Me.Button27.TabIndex = 177
        Me.Button27.Text = "Delete"
        Me.Button27.UseVisualStyleBackColor = False
        '
        'TextBox7
        '
        Me.TextBox7.Location = New System.Drawing.Point(15, 21)
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.Size = New System.Drawing.Size(56, 22)
        Me.TextBox7.TabIndex = 180
        Me.TextBox7.Text = "53"
        Me.TextBox7.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox8
        '
        Me.TextBox8.Location = New System.Drawing.Point(18, 20)
        Me.TextBox8.Name = "TextBox8"
        Me.TextBox8.Size = New System.Drawing.Size(56, 22)
        Me.TextBox8.TabIndex = 179
        Me.TextBox8.Text = "39"
        Me.TextBox8.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Snow
        Me.Button1.ForeColor = System.Drawing.Color.Red
        Me.Button1.Location = New System.Drawing.Point(81, 82)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(69, 26)
        Me.Button1.TabIndex = 181
        Me.Button1.Text = "+12 v"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'TextBox9
        '
        Me.TextBox9.Location = New System.Drawing.Point(1211, 43)
        Me.TextBox9.Name = "TextBox9"
        Me.TextBox9.Size = New System.Drawing.Size(39, 22)
        Me.TextBox9.TabIndex = 183
        Me.TextBox9.Text = "2"
        Me.TextBox9.Visible = False
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Snow
        Me.Button2.Location = New System.Drawing.Point(75, 59)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(55, 27)
        Me.Button2.TabIndex = 184
        Me.Button2.Text = "Save"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.Snow
        Me.Button5.Location = New System.Drawing.Point(6, 53)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(69, 26)
        Me.Button5.TabIndex = 185
        Me.Button5.Text = "0 - 5 v"
        Me.Button5.UseVisualStyleBackColor = False
        '
        'Button7
        '
        Me.Button7.BackColor = System.Drawing.Color.Snow
        Me.Button7.ForeColor = System.Drawing.Color.Green
        Me.Button7.Location = New System.Drawing.Point(81, 21)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(69, 26)
        Me.Button7.TabIndex = 186
        Me.Button7.Text = "Gnd"
        Me.Button7.UseVisualStyleBackColor = False
        '
        'Button15
        '
        Me.Button15.BackColor = System.Drawing.Color.Snow
        Me.Button15.ForeColor = System.Drawing.Color.DarkRed
        Me.Button15.Location = New System.Drawing.Point(81, 53)
        Me.Button15.Name = "Button15"
        Me.Button15.Size = New System.Drawing.Size(69, 26)
        Me.Button15.TabIndex = 187
        Me.Button15.Text = "+5 v"
        Me.Button15.UseVisualStyleBackColor = False
        '
        'TextBox24
        '
        Me.TextBox24.Location = New System.Drawing.Point(1211, 17)
        Me.TextBox24.Name = "TextBox24"
        Me.TextBox24.Size = New System.Drawing.Size(119, 22)
        Me.TextBox24.TabIndex = 107
        Me.TextBox24.Text = "10"
        Me.TextBox24.Visible = False
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.Snow
        Me.Button3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Button3.Location = New System.Drawing.Point(6, 82)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(69, 26)
        Me.Button3.TabIndex = 188
        Me.Button3.Text = "0 - 12 v"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TextBox7)
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox1.Location = New System.Drawing.Point(175, 13)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(95, 53)
        Me.GroupBox1.TabIndex = 189
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Width"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.TextBox8)
        Me.GroupBox2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox2.Location = New System.Drawing.Point(175, 72)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(95, 53)
        Me.GroupBox2.TabIndex = 190
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Height"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Button4)
        Me.GroupBox3.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox3.Location = New System.Drawing.Point(476, 64)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(272, 62)
        Me.GroupBox3.TabIndex = 193
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Side of Stripboard"
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.Snow
        Me.Button4.Location = New System.Drawing.Point(6, 21)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(260, 27)
        Me.Button4.TabIndex = 191
        Me.Button4.Text = "Front"
        Me.Button4.UseVisualStyleBackColor = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Button7)
        Me.GroupBox4.Controls.Add(Me.Button5)
        Me.GroupBox4.Controls.Add(Me.Button3)
        Me.GroupBox4.Controls.Add(Me.Button15)
        Me.GroupBox4.Controls.Add(Me.Button1)
        Me.GroupBox4.Controls.Add(Me.Button42)
        Me.GroupBox4.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox4.Location = New System.Drawing.Point(294, 12)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(162, 114)
        Me.GroupBox4.TabIndex = 194
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Jumpers"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.ComboBox7)
        Me.GroupBox5.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox5.Location = New System.Drawing.Point(476, 12)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(272, 48)
        Me.GroupBox5.TabIndex = 195
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Component Footprints"
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.Button2)
        Me.GroupBox6.Controls.Add(Me.Button43)
        Me.GroupBox6.Controls.Add(Me.Button40)
        Me.GroupBox6.Controls.Add(Me.Button8)
        Me.GroupBox6.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox6.Location = New System.Drawing.Point(14, 13)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(145, 112)
        Me.GroupBox6.TabIndex = 196
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Stripboard"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.Button9)
        Me.GroupBox7.Controls.Add(Me.Button6)
        Me.GroupBox7.Controls.Add(Me.Button27)
        Me.GroupBox7.Controls.Add(Me.Button41)
        Me.GroupBox7.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox7.Location = New System.Drawing.Point(759, 17)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(228, 104)
        Me.GroupBox7.TabIndex = 197
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Misc"
        '
        'Button9
        '
        Me.Button9.BackColor = System.Drawing.Color.Snow
        Me.Button9.Location = New System.Drawing.Point(118, 55)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(106, 27)
        Me.Button9.TabIndex = 179
        Me.Button9.Text = "Grid"
        Me.Button9.UseVisualStyleBackColor = False
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.Snow
        Me.Button6.Location = New System.Drawing.Point(118, 21)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(106, 25)
        Me.Button6.TabIndex = 178
        Me.Button6.Text = "Twixt Break"
        Me.Button6.UseVisualStyleBackColor = False
        '
        'TextBox22
        '
        Me.TextBox22.Location = New System.Drawing.Point(1082, 74)
        Me.TextBox22.Name = "TextBox22"
        Me.TextBox22.Size = New System.Drawing.Size(43, 22)
        Me.TextBox22.TabIndex = 168
        Me.TextBox22.Visible = False
        '
        'TextBox18
        '
        Me.TextBox18.Location = New System.Drawing.Point(1137, 17)
        Me.TextBox18.Name = "TextBox18"
        Me.TextBox18.Size = New System.Drawing.Size(43, 22)
        Me.TextBox18.TabIndex = 165
        Me.TextBox18.Visible = False
        '
        'TextBox23
        '
        Me.TextBox23.Location = New System.Drawing.Point(1345, 17)
        Me.TextBox23.Name = "TextBox23"
        Me.TextBox23.Size = New System.Drawing.Size(31, 22)
        Me.TextBox23.TabIndex = 105
        Me.TextBox23.Text = "7"
        Me.TextBox23.Visible = False
        '
        'TextBox32
        '
        Me.TextBox32.Location = New System.Drawing.Point(1211, 77)
        Me.TextBox32.Name = "TextBox32"
        Me.TextBox32.Size = New System.Drawing.Size(29, 22)
        Me.TextBox32.TabIndex = 146
        Me.TextBox32.Text = "30"
        Me.TextBox32.Visible = False
        '
        'ListBox20
        '
        Me.ListBox20.BackColor = System.Drawing.Color.AntiqueWhite
        Me.ListBox20.ItemHeight = 16
        Me.ListBox20.Location = New System.Drawing.Point(1246, 129)
        Me.ListBox20.Name = "ListBox20"
        Me.ListBox20.Size = New System.Drawing.Size(275, 468)
        Me.ListBox20.TabIndex = 85
        '
        'ListBox19
        '
        Me.ListBox19.BackColor = System.Drawing.Color.AntiqueWhite
        Me.ListBox19.ItemHeight = 16
        Me.ListBox19.Location = New System.Drawing.Point(1281, 160)
        Me.ListBox19.Name = "ListBox19"
        Me.ListBox19.Size = New System.Drawing.Size(138, 100)
        Me.ListBox19.TabIndex = 171
        Me.ListBox19.Visible = False
        '
        'ListBox2
        '
        Me.ListBox2.BackColor = System.Drawing.Color.AntiqueWhite
        Me.ListBox2.ItemHeight = 16
        Me.ListBox2.Location = New System.Drawing.Point(1264, 129)
        Me.ListBox2.Name = "ListBox2"
        Me.ListBox2.Size = New System.Drawing.Size(138, 148)
        Me.ListBox2.TabIndex = 4
        '
        'TextBox19
        '
        Me.TextBox19.Location = New System.Drawing.Point(1246, 74)
        Me.TextBox19.Name = "TextBox19"
        Me.TextBox19.Size = New System.Drawing.Size(29, 22)
        Me.TextBox19.TabIndex = 84
        Me.TextBox19.Text = "1"
        Me.TextBox19.Visible = False
        '
        'TextBox33
        '
        Me.TextBox33.Location = New System.Drawing.Point(1264, 45)
        Me.TextBox33.Name = "TextBox33"
        Me.TextBox33.Size = New System.Drawing.Size(29, 22)
        Me.TextBox33.TabIndex = 147
        Me.TextBox33.Text = "1"
        Me.TextBox33.Visible = False
        '
        'TextBox10
        '
        Me.TextBox10.Location = New System.Drawing.Point(1301, 42)
        Me.TextBox10.Name = "TextBox10"
        Me.TextBox10.Size = New System.Drawing.Size(29, 22)
        Me.TextBox10.TabIndex = 198
        Me.TextBox10.Text = "1"
        Me.TextBox10.Visible = False
        '
        'TextBox17
        '
        Me.TextBox17.Location = New System.Drawing.Point(1281, 77)
        Me.TextBox17.Name = "TextBox17"
        Me.TextBox17.Size = New System.Drawing.Size(29, 22)
        Me.TextBox17.TabIndex = 199
        Me.TextBox17.Text = "1"
        Me.TextBox17.Visible = False
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.BackColor = System.Drawing.Color.Linen
        Me.ClientSize = New System.Drawing.Size(1200, 966)
        Me.Controls.Add(Me.TextBox17)
        Me.Controls.Add(Me.TextBox10)
        Me.Controls.Add(Me.GroupBox7)
        Me.Controls.Add(Me.GroupBox6)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TextBox9)
        Me.Controls.Add(Me.ListBox27)
        Me.Controls.Add(Me.ListBox25)
        Me.Controls.Add(Me.ListBox23)
        Me.Controls.Add(Me.ListBox19)
        Me.Controls.Add(Me.ListBox28)
        Me.Controls.Add(Me.TextBox22)
        Me.Controls.Add(Me.TextBox21)
        Me.Controls.Add(Me.TextBox20)
        Me.Controls.Add(Me.TextBox18)
        Me.Controls.Add(Me.TextBox15)
        Me.Controls.Add(Me.TextBox14)
        Me.Controls.Add(Me.TextBox35)
        Me.Controls.Add(Me.TextBox33)
        Me.Controls.Add(Me.TextBox32)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.TextBox26)
        Me.Controls.Add(Me.TextBox27)
        Me.Controls.Add(Me.ListBox26)
        Me.Controls.Add(Me.TextBox24)
        Me.Controls.Add(Me.TextBox23)
        Me.Controls.Add(Me.listbox5)
        Me.Controls.Add(Me.ListBox24)
        Me.Controls.Add(Me.ListBox21)
        Me.Controls.Add(Me.ListBox20)
        Me.Controls.Add(Me.TextBox19)
        Me.Controls.Add(Me.ListBox10)
        Me.Controls.Add(Me.ListBox12)
        Me.Controls.Add(Me.ListBox11)
        Me.Controls.Add(Me.ListBox9)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.ListBox4)
        Me.Controls.Add(Me.ListBox3)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.ListBox2)
        Me.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Form1"
        Me.ShowIcon = False
        Me.Text = "Clanking Replicator Stripboard Designer"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region


    'Dim bmp As New Bitmap(PictureBox1.Image)
    'Dim graphics As Graphics = Graphics.FromImage(bmp)

    Dim x_md, y_md As Int16 'cordinats of the line
    Dim x_mm, y_mm As Int16     'cordinats of the line
    Dim myPen = New Pen(Color.Red, 1)
    '*************************************************************************
    '*************************************************************************
    '*************************************************************************
    '*************************************************************************

    Dim XBoardLimit As Single = 1500
    Dim YBoardLimit As Single = 1000
    Dim PinSpacing As Single = 25.4 '* 10

    Dim ToolFlag As String = ""

    Dim OffsetX, OffsetY As Double
    Dim ShowIndex As Integer = 0

    Dim MaxMin(1, 1) As Integer

    '*************************************************************************
    '*************************************************************************

    Dim BasePath As String = CurDir()

    '*************************************************************************
    '*************************************************************************

    Dim StepCounter As Long = 0
    Dim ExtrudeCounter As Long = 0
    Dim NoExtrudeCounter As Long = 0

    Dim ExtruderStatus As Boolean

    Dim absolute_x As Integer
    Dim absolute_y As Integer
    Dim LastInputString As String

    '*************************************************************************
    '*************************************************************************
    '*************************************************************************
    '*************************************************************************

    Dim BottomLayer As Short
    Dim TopLayer As Short

    Dim instance As Bitmap '= New Bitmap(250, 250)
    Dim pic1x As Integer ' = e.X
    Dim pic1y As Integer ' = e.Y


    Dim lastcolor As Color = Color.White

    Dim ResetToFirstPointFlag As Boolean
    Dim DisplayScale As Single
    Dim GridBoundary As Integer = 2400
    Dim PenWidth As Single = 1

    Dim Triangles(1, 3, 3) As Double
    Dim TriangleSegments(300000, 3, 2, 3) As Double
    Dim PlanarCoefficients(300000, 4) As Double

    Dim SliceCoefficients(4) As Double

    Dim FacetCounter As Integer = 0
    Dim CentreX As Integer = Me.Size.Width / 2
    Dim CentreY As Integer = Me.Size.Height / 2

    Dim A(10, 10) As Double
    Dim B(10, 11) As Double
    Dim X(10) As Double
    Dim SliceBoundary(100000, 2, 3) As Double
    Dim SliceLimit As Short
    Dim ClockwisePerimeter(15000, 2, 3) As Double
    Dim ClockwisePerimeterTag(15000) As Boolean
    Dim ClockwisePerimeterSurface(15000) As Long
    Dim ClockwiseLimit As Long

    Dim LinkedClockwisePerimeter(15000, 2, 3) As Double
    Dim LinkedClockwiseLimit As Long
    Dim LinkedClockwisePerimeterSurface(15000) As Long
    Dim LinkedClockwisePerimeterLoopID(15000) As Short

    Dim LinkedLoopLimit As Short

    Dim ClockwiseLoop(15000) As Short
    Dim ConsolidatedPerimeterLoop(15000, 2, 3) As Double
    Dim ConsolidatedPerimeterLoopLimit As Long
    Dim ConsolidatedPerimeterSurface(15000) As Long
    Dim ConsolidatedPerimeterLoopID(15000) As Short

    Dim MkIIPerimeterLoop(15000, 2, 3) As Double
    Dim MkIIPerimeterLoopLimit As Long
    Dim MkIIPerimeterSurface(15000) As Long
    Dim MkIIPerimeterLoopID(15000) As Short

    Dim Grid(5000, 5000) As Byte
    Dim ScratchGrid(5000, 5000) As Byte
    Dim Last_Grid(5000, 5000) As Byte
    Dim Infill_Grid(5000, 5000) As Byte
    Dim Boundaries(2, 3) As Double
    Dim LeftOrRight As String = "right"

    Dim StraightenXLast As Integer
    Dim StraightenYLast As Integer

    Dim PCBGrid(5000, 5000) As Short
    Dim PCBGridSteps(5000, 5000) As Short
    Dim DevelopSliceActiveFlag(20000) As Boolean

    Dim topLoop(100, 4) As segment
    Dim topLoopIndex As Integer = 0

    Dim bottomLoop(100, 4) As segment
    Dim bottomLoopIndex As Integer = 0

    Structure vector
        Dim x As Double
        Dim y As Double
        Dim z As Double
        Dim rho As Double
        Dim theta As Double
    End Structure
    Structure segment
        Dim P1 As vector
        Dim P2 As vector
    End Structure

    Structure Point
        Dim LineNumber As Integer
        Dim X As Integer
        Dim Y As Integer
    End Structure

    Structure Path
        Dim Point1 As Point
        Dim Point2 As Point
        Dim Active As Boolean
        Dim Action As String
    End Structure

    Dim Paths(1000) As Path
    Dim SavePaths(1000) As Path

    Dim HitCount As Integer = 0


    Private Sub CylindricalToCartesian(ByVal Rho As Double, ByVal ThetaInDegrees As Double, ByRef x As Double, ByRef y As Double)

        '   Convert Theta from degrees to radians
        Dim ThetaInRadians As Double = (2 * Math.PI / 180) * ThetaInDegrees
        x = Rho * Math.Cos(ThetaInRadians)
        y = Rho * Math.Sin(ThetaInRadians)
    End Sub

    Private Sub clearscreen(ByVal gr As Graphics)
        gr.Clear(Me.BackColor)
    End Sub
    Public Sub WaitOneSeconds()
        Dim Start, Finish As Double

        Start = Microsoft.VisualBasic.DateAndTime.Timer
        Finish = Start + 0.05 ' Set end time for 5-second duration.
        Do While Microsoft.VisualBasic.DateAndTime.Timer < Finish
            ' Do other processing while waiting for 5 seconds to elapse.
        Loop
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim mypath1 As String = Mid(CurDir, 1, CurDir.Length - 8)
        Dim myname As String
        '**********************************************************************************
        '**********************************************************************************
        '**********************************************************************************
        '**********************************************************************************
        GroupBox4.Visible = False
        GroupBox5.Visible = False
        GroupBox3.Visible = False
        GroupBox7.Visible = False
        'GroupBox1.Visible = False
        'GroupBox2.Visible = False
        Button8.Visible = False
        Button2.Visible = False
        '**********************************************************************************
        '**********************************************************************************
        '**********************************************************************************
        '**********************************************************************************
        'TextBox6.Text = "0.102"
        Dim bmp01 As New Bitmap(PictureBox1.Width, PictureBox1.Height)
        PictureBox1.Image = bmp01
        ''listbox9.Visible = False
        'ListBox10.Visible = False
        'ListBox11.Visible = False
    End Sub

    Public Function CrossProduct(ByVal A1 As Double, ByVal A2 As Double, ByVal A3 As Double, _
                                 ByVal B1 As Double, ByVal B2 As Double, ByVal B3 As Double, _
                                 ByRef C1 As Double, ByRef C2 As Double, ByRef C3 As Double) As Boolean
        C1 = A2 * B3 - A3 * B2
        C2 = A3 * B1 - A1 * B3
        C3 = A1 * B2 - A2 * B1
        If C1 = 0 And C2 = 0 And C3 = 0 Then
            CrossProduct = False
        Else
            CrossProduct = True
        End If
    End Function

    Public Function CramersRule(ByVal a1 As Double, ByVal b1 As Double, ByVal c1 As Double, _
                             ByVal a2 As Double, ByVal b2 As Double, ByVal c2 As Double, _
                             ByRef x1 As Double, ByRef x2 As Double) As Boolean


        Dim x1Denominator As Double = (b2 * a1 - b1 * a2)
        Dim x2Denominator As Double = (a2 * b1 - a1 * b2)

        If x1Denominator = 0 Or x2Denominator = 0 Then
            CramersRule = False
            x1 = 0
            x2 = 0
        Else
            CramersRule = True
            x1 = (b2 * c1 - b1 * c2) / x1Denominator
            x2 = (a2 * c1 - a1 * c2) / x2Denominator
        End If
    End Function


    Public Function Equivalent(ByVal A As Double, ByVal B As Double, ByVal Tolerance As Double) As Boolean
        If Math.Abs(A - B) < Tolerance Then
            Equivalent = True
        Else
            Equivalent = False
        End If
    End Function

    Public Function Done() As Short
        Dim i As Long '= 0
        Done = 0
        For i = 1 To ClockwiseLimit
            If ClockwisePerimeterTag(i) = True Then
                Done = i
                Exit Function
            End If
        Next
    End Function

    Public Function SameDirection(ByVal A1 As Double, ByVal A2 As Double, ByVal A3 As Double, ByVal B1 As Double, ByVal B2 As Double, ByVal B3 As Double, ByVal Tolerance As Double) As Boolean
        '     normalise both vectors

        Dim rho1 As Double
        Dim rho2 As Double

        'Dim i As Short

        rho1 = (A1 * A1 + A2 * A2 + A3 * A3) ^ 0.5
        rho2 = (B1 * B1 + B2 * B2 + B3 * B3) ^ 0.5
        If Math.Abs(rho1) < Tolerance Or Math.Abs(rho2) < Tolerance Then
            SameDirection = False
            Exit Function
        End If

        '     normalise the vectors

        A1 = A1 / rho1
        A2 = A2 / rho1
        A3 = A3 / rho1

        B1 = B1 / rho2
        B2 = B2 / rho2
        B3 = B3 / rho2

        Dim DotProductAB As Double = A1 * B1 + A2 * B2 + A3 * B3

        If Math.Abs(DotProductAB - 1) < Tolerance Then
            SameDirection = True
        Else
            SameDirection = False
        End If

    End Function


    Public Function CheckPathList(ByVal XCoordinate As Integer, ByVal YCoordinate As Integer) As Boolean
        Dim sp As Object
        Dim j As Integer
        Dim XCoordinate1, YCoordinate1 As Integer
        If ListBox2.Items.Count > 0 Then
            For j = ListBox2.Items.Count - 1 To 0 Step -1
                sp = Split(ListBox2.Items.Item(j), Chr(9))
                XCoordinate1 = CInt(sp(0))
                YCoordinate1 = CInt(sp(1))
                If XCoordinate = XCoordinate1 And YCoordinate = YCoordinate1 Then
                    CheckPathList = True
                    Exit Function
                End If
            Next
        End If
        CheckPathList = False
    End Function

    Private Sub PictureBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove
        Exit Sub
        TextBox1.Text = e.X
        TextBox2.Text = e.Y
        TextBox1.Refresh()
        TextBox2.Refresh()

        'Dim x As Integer = e.X
        'Dim y As Integer = e.Y
        pic1x = e.X
        pic1y = e.Y

        'Dim instance As Bitmap = New Bitmap("..\slice 01.bmp")
        Dim returnValue As Color


        returnValue = instance.GetPixel(pic1x, pic1y)

        If returnValue = lastcolor Then

            TextBox3.Text = "same"
        Else
            TextBox3.Text = "changed"
            ListBox1.Items.Add(Str(pic1x) & Chr(9) & Str(pic1y))

            lastcolor = returnValue
        End If
        TextBox3.Refresh()

    End Sub

    Private Sub GetSimplePattern(ByVal WholePattern As String)
        Dim X As Integer
        Dim Y As Integer
        Dim Pattern As String = Mid(WholePattern, 5, 1)
        Dim Size As Integer = CInt(Mid(WholePattern, 1, 4))
        Select Case Pattern
            Case "A"
                X = -1 * Size
                Y = -1 * Size
            Case "B"
                X = -1 * Size
                Y = 0 * Size
            Case "C"
                X = -1 * Size
                Y = 1 * Size
            Case "D"
                X = 0 * Size
                Y = -1 * Size
            Case "E"
                X = 0 * Size
                Y = 1 * Size
            Case "F"
                X = 1 * Size
                Y = -1 * Size
            Case "G"
                X = 1 * Size
                Y = 0 * Size
            Case "H"
                X = 1 * Size
                Y = 1 * Size
        End Select
        ''listbox9.Items.Add(Chr(9) & Chr(9) & Chr(9) & "<AA1 point x=" & Chr(34) & LTrim(RTrim(Str(X))) & Chr(34) & " y=" & Chr(34) & LTrim(RTrim(Str(Y))) & Chr(34) & "/>")
        X = X + StraightenXLast
        Y = Y + StraightenYLast
        ''listbox9.Items.Add(Chr(9) & Chr(9) & Chr(9) & "<point x=" & Chr(34) & LTrim(RTrim(Str(X))) & Chr(34) & " y=" & Chr(34) & LTrim(RTrim(Str(Y))) & Chr(34) & "/>")
        PrintLine(101, Chr(9) & Chr(9) & Chr(9) & "<point x=" & Chr(34) & LTrim(RTrim(Str(X))) & Chr(34) & " y=" & Chr(34) & LTrim(RTrim(Str(Y))) & Chr(34) & "/>")
        StraightenXLast = X
        StraightenYLast = Y
    End Sub


    Private Sub XY(ByVal TextString As String, ByRef xx As Integer, ByRef yy As Integer)

        Dim sp As Object
        sp = Split(TextString, Chr(34))
        Try
            xx = CDbl(sp(1))
            yy = CDbl(sp(3))
        Catch ex As Exception
            ListBox19.Items.Add("exiting")
            Exit Sub
        End Try


    End Sub

    Public Function PlanarIntercept(ByVal X1 As Double, ByVal Y1 As Double, ByVal Z1 As Double, _
                                   ByVal X2 As Double, ByVal Y2 As Double, ByVal Z2 As Double, ByRef T As Double) As Short
        Dim Denominator As Double
        Denominator = SliceCoefficients(1) * (X2 - X1) + SliceCoefficients(2) * (Y2 - Y1) + SliceCoefficients(3) * (Z2 - Z1)

        T = SliceCoefficients(1) * X1 + SliceCoefficients(2) * Y1 + SliceCoefficients(3) * Z1 + SliceCoefficients(4)
        If Denominator = 0 Then
            PlanarIntercept = 0
        Else
            T = -T / Denominator
            If T >= 0 And T < 1 Then
                PlanarIntercept = 1
            Else
                PlanarIntercept = -1
            End If
        End If
        'ListBox1.Items.Add("      T = " & Format(T, "00.00") & "   planarintercept = " & PlanarIntercept)
    End Function

    Private Function FindClosest(ByVal x As Integer, ByVal y As Integer) As Integer
        '################################################################################################
        '################################################################################################

        Dim spi0, spi1 As Object
        Dim LocationX, LocationY As Integer
        Dim j As Integer

        Dim RhoClosest As Double = 10000000000.0
        Dim IndexClosest, XClosest, YClosest As Integer
        Dim RhoScratch As Double

        'Dim TruthFlag As Boolean

        '################################################################################################
        '################################################################################################

        '   find the closest point to the origin

        For j = 0 To ListBox23.Items.Count - 2 Step 2
            If DevelopSliceActiveFlag(j) = True Then



                spi0 = Split(ListBox23.Items.Item(j), "|")
                spi1 = Split(ListBox23.Items.Item(j + 1), "|")

                LocationX = CInt(spi0(0))
                LocationY = CInt(spi0(1))

                ListBox24.Items.Add("point = " & Chr(9) & Str(j) & Chr(9) & Str(LocationX) & Chr(9) & Str(LocationY))

                RhoScratch = ((LocationX) * (LocationX) + (LocationY) * (LocationY)) ^ 0.5

                If RhoScratch < RhoClosest Then
                    RhoClosest = RhoScratch
                    XClosest = LocationX
                    YClosest = LocationY
                    IndexClosest = j
                    'ListBox24.Items.Add("closest point = " & Chr(9) & Str(IndexClosest) & Chr(9) & Format(RhoClosest, "0000.0") & Str(XClosest) & Chr(9) & Str(YClosest))
                End If

            End If
        Next
        DevelopSliceActiveFlag(IndexClosest) = False
        ListBox24.Items.Add("closest point = " & Chr(9) & Str(IndexClosest) & Chr(9) & Format(RhoClosest, "0000.0") & Str(XClosest) & Chr(9) & Str(YClosest))

        FindClosest = IndexClosest

        '################################################################################################
        '################################################################################################
    End Function

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim bmp As New Bitmap(PictureBox1.Image)
        Dim graphics As Graphics = graphics.FromImage(bmp)

        GroupBox4.Visible = False
        GroupBox5.Visible = False
        GroupBox3.Visible = False
        GroupBox7.Visible = False
        'GroupBox1.Visible = False
        'GroupBox2.Visible = False
        Button8.Visible = False
        Button2.Visible = False

        'Dim ii, iii As Integer
        graphics.Clear(Color.Ivory)

        PictureBox1.Image = bmp
        graphics.Dispose()
        PictureBox1.Refresh()

    End Sub



    Public Function Distance(ByVal x1 As Integer, ByVal y1 As Integer, ByRef x2 As Integer, ByVal y2 As Integer) As Double
        Distance = ((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1)) ^ 0.5
    End Function


    Private Sub ComboBox7_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox7.Click
        ToolFlag = "footprint"
        PenWidth = CInt(TextBox9.Text)
    End Sub

    Private Function NewDotProduct(ByVal a As vector, ByVal b As vector) As Double
        Dim i As Integer
        NewDotProduct = (a.x * b.x + a.y * b.y + a.z * b.z)
    End Function

    Private Function Normalise(ByVal v As vector) As vector
        Dim i As Integer
        Normalise.rho = (v.x * v.x + v.y * v.y) ^ 0.5
        Normalise.x = v.x / Normalise.rho
        Normalise.y = v.y / Normalise.rho
        Normalise.z = v.z / Normalise.rho
    End Function

    Public Function NewCrossProduct(ByVal a As vector, ByVal b As vector) As vector
        NewCrossProduct.x = a.y * b.z - a.z * b.y
        NewCrossProduct.y = a.z * b.x - a.x * b.z
        NewCrossProduct.z = a.x * b.y - a.y * b.x
    End Function

    Public Function NewCramersRule(ByVal S1 As segment, ByVal S2 As segment, _
                             ByRef T1 As Double, ByRef T2 As Double) As Boolean
        'Public Function NewCramersRule(ByVal a1 As Double, ByVal b1 As Double, ByVal c1 As Double, _
        '                         ByVal a2 As Double, ByVal b2 As Double, ByVal c2 As Double, _
        '                         ByRef x1 As Double, ByRef x2 As Double) As Boolean

        Dim a1 As Double
        Dim a2 As Double
        Dim a3 As Double

        Dim b1 As Double
        Dim b2 As Double
        Dim b3 As Double

        Dim c1 As Double
        Dim c2 As Double
        Dim c3 As Double

        'a1 = (XStep1 + scanX) - (XStep + scanX)
        a1 = S1.P2.x - S1.P1.x
        'b1 = -(ConsolidatedPerimeterLoop(i, 2, 1) - ConsolidatedPerimeterLoop(i, 1, 1))
        b1 = -(S2.P2.x - S2.P1.x)
        'c1 = ConsolidatedPerimeterLoop(i, 1, 1) - (XStep + scanX)
        c1 = S2.P1.x - S1.P1.x

        'a2 = (ZStep1) - (ZStep)
        a2 = S1.P2.y - S1.P1.y
        'b2 = -(ConsolidatedPerimeterLoop(i, 2, 3) - ConsolidatedPerimeterLoop(i, 1, 3))
        b2 = -(S2.P2.y - S2.P1.y)
        'c2 = ConsolidatedPerimeterLoop(i, 1, 3) - (ZStep)
        c2 = S2.P1.y - S1.P1.y

        Dim T1Denominator As Double = (b2 * a1 - b1 * a2)
        Dim T2Denominator As Double = (a2 * b1 - a1 * b2)

        If T1Denominator = 0 Or T2Denominator = 0 Then
            '   parallel
            NewCramersRule = False
            T1 = 0
            T2 = 0
        Else
            NewCramersRule = True
            T1 = (b2 * c1 - b1 * c2) / T1Denominator
            T2 = (a2 * c1 - a1 * c2) / T2Denominator
        End If
    End Function
    Private Function NewCartesianToCylindrical(ByVal V As vector) As vector
        Dim scratchVector As vector = V
        V = Normalise(V)

        NewCartesianToCylindrical.rho = V.rho
        NewCartesianToCylindrical.x = V.x
        NewCartesianToCylindrical.y = V.y
        NewCartesianToCylindrical.z = V.z

        Select Case V.x
            Case Is > 0
                NewCartesianToCylindrical.theta = Math.Atan(V.y / V.x)
            Case Is = 0
                Select Case V.y
                    Case Is > 0
                        NewCartesianToCylindrical.theta = V.rho / 2
                    Case Is < 0
                        NewCartesianToCylindrical.theta = 3 * V.rho / 2
                End Select
            Case Is < 0
                NewCartesianToCylindrical.rho = V.rho + Math.Atan(V.y / V.x)
        End Select
        'ListBox20.Items.Add(Chr(9) & "cart2cylind " & Chr(9) & Format(NewCartesianToCylindrical.theta, "000.00") & Chr(9) & Format(NewCartesianToCylindrical.rho, "000.00"))

    End Function
    Private Function NewerCylindricalToCartesian(ByVal V As vector) As vector
        'V.rho = (V.x * V.x + V.y * V.y) ^ 0.5
        NewerCylindricalToCartesian.x = V.rho * Math.Cos(V.theta)
        NewerCylindricalToCartesian.y = V.rho * Math.Sin(V.theta)
        NewerCylindricalToCartesian.z = V.z
        NewerCylindricalToCartesian.theta = 0
        NewerCylindricalToCartesian.rho = V.rho
        'ListBox20.Items.Add(Chr(9) & "cylind2cart" & Chr(9) & Format(NewerCylindricalToCartesian.x, "000.00") & Chr(9) & Format(NewerCylindricalToCartesian.z, "000.00"))

    End Function
    Private Function NewCylindricalToCartesian(ByVal Rho As Double, ByVal ThetaInDegrees As Double, ByVal V As vector) As vector
        '   Convert Theta from degrees to radians
        Dim ThetaInRadians As Double = (Math.PI / 180) * ThetaInDegrees

        NewCylindricalToCartesian.x = Rho * Math.Cos(ThetaInRadians)
        NewCylindricalToCartesian.y = Rho * Math.Sin(ThetaInRadians)
        NewCylindricalToCartesian.rho = Rho
        NewCylindricalToCartesian.theta = ThetaInDegrees

    End Function

    Private Function AddVectors(ByVal A As vector, ByVal B As vector) As vector
        AddVectors.x = A.x + B.x
        AddVectors.y = A.y + B.y
        'AddVectors.rho = A.rho + B.rho
        'AddVectors.theta = A.theta + B.theta
    End Function
    Private Function SubtractVectors(ByVal A As vector, ByVal B As vector) As vector
        SubtractVectors.x = A.x - B.x
        SubtractVectors.y = A.y - B.y
        'SubtractVectors.rho = A.rho - B.rho
        'SubtractVectors.theta = A.theta - B.theta
    End Function
    Private Sub UpdatedCylindricalToCartesian(ByVal Rho As Double, ByVal ThetaInDegrees As Double, ByRef x As Double, ByRef y As Double)

        '   Convert Theta from degrees to radians
        Dim ThetaInRadians As Double = (Math.PI / 180) * ThetaInDegrees
        x = Rho * Math.Cos(ThetaInRadians)
        y = Rho * Math.Sin(ThetaInRadians)
    End Sub


    Sub makeArc(ByVal direction As String, ByVal Offset As vector, ByVal Radius As Double, ByVal startTheta As Double, ByVal stopTheta As Double, ByVal bmp As Bitmap)
        '********************************************************************
        '********************************************************************
        '********************************************************************
        '********************************************************************

        'Dim bmp As New Bitmap(PictureBox1.Image)
        'Dim graphics As Graphics = graphics.FromImage(bmp)

        'graphics.Clear(Color.White)

        'PictureBox1.Image = bmp
        'graphics.Dispose()
        'PictureBox1.Refresh()

        PenWidth = CSng(TextBox9.Text)

        Dim XDisplay1 As Single
        Dim YDisplay1 As Single

        Dim XDisplay2 As Single
        Dim YDisplay2 As Single

        Dim BoundariesFlag As Boolean = False

        Dim m As Integer

        '********************************************************************
        '********************************************************************
        PenWidth = CSng(TextBox9.Text)

        DisplayScale = CSng(TextBox35.Text)
        Dim Degrees As Double

        '****************************************************************************************
        '****************************************************************************************
        '****************************************************************************************
        '****************************************************************************************

        Dim PegSize As Double = 0.3
        Dim DegreeStep1 As Double = 10

        '****************************************************************************************
        '   Outer boundary
        '****************************************************************************************
        '****************************************************************************************

        Dim V1 As vector

        V1.theta = startTheta * Math.PI / 180
        V1.rho = Radius
        V1 = NewerCylindricalToCartesian(V1)

        V1 = AddVectors(V1, Offset)

        Dim V2 As vector

        'Dim stopTheta As Double = startTheta - endTheta

        Dim steps As Double = Math.Abs(startTheta - stopTheta) / DegreeStep1
        steps = Math.Truncate(steps) + 1
        Dim DegreeStep As Double = Math.Abs(startTheta - stopTheta) / steps

        If direction = "+" Then
            DegreeStep = Math.Sqrt(DegreeStep * DegreeStep)
        ElseIf direction = "-" Then
            DegreeStep = -1 * Math.Sqrt(DegreeStep * DegreeStep)
            'stopTheta = startTheta - deltaTheta
        End If
        Dim arcDegrees As Double
        Dim lastArcDegrees As Double

        lastArcDegrees = startTheta
        Dim evenodd As Boolean = False

        ListBox20.Items.Add("Cap")

        ListBox20.Items.Add("angles = " & Chr(9) & Format(startTheta, "000.00") & Chr(9) & Format(stopTheta, "000.00"))
        ListBox20.Items.Add("degree step = " & Chr(9) & Format(DegreeStep, "000.000"))

        ListBox1.Items.Clear()
        For Degrees = startTheta To stopTheta + 0.01 Step DegreeStep

            'PrintLine(101, "degrees = " & Format(Degrees, "000.0"))

            'If DegreeStep < 0 And Degrees < 0 Then
            'arcDegrees = 360 - Degrees
            'End If

            V2.rho = Radius

            'Call UpdatedCylindricalToCartesian(Radius, Degrees, V2.x, V2.y)

            V2.theta = Degrees * Math.PI / 180

            V2 = NewerCylindricalToCartesian(V2)

            V2 = AddVectors(V2, Offset)
            'V2.x = V2.x + Offset.x
            'V2.y = V2.y + Offset.y

            XDisplay1 = (V1.x) / DisplayScale
            YDisplay1 = (V1.y) / DisplayScale
            XDisplay2 = (V2.x) / DisplayScale
            YDisplay2 = (V2.y) / DisplayScale

            ListBox20.Items.Add(Chr(9) & Chr(9) & Format(Degrees, "000.00") & Chr(9) & Format(V1.x / 10, "000.00") & Chr(9) & Format(V1.y / 10, "000.00") & _
                                                         Chr(9) & Format(V2.x / 10, "000.00") & Chr(9) & Format(V1.y / 10, "000.00"))

            ListBox20.Items.Add("xx" & Chr(9) & Chr(9) & Format(V2.x / 10, "000.00") & Chr(9) & Format(V1.y / 10, "000.00"))
            'ListBox27.Items.Add("zz" & Chr(9) & Chr(9) & Format(V2.x / 10, "000.00") & Chr(9) & Format(V1.y / 10, "000.00"))

            ''PrintLine(101, "zz" & Chr(9) & Chr(9) & Format(V2.x / 10, "000.00") & Chr(9) & Format(V1.y / 10, "000.00"))

            If direction = "+" Then
                'Call NewDrawPurpleLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Purple", bmp)
                evenodd = True
            Else
                'Call NewDrawOrchidLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Orchid", bmp)
                evenodd = False
            End If
            'PrintLine(101, Chr(9) & Chr(9) & Chr(9) & "<point x=" & Chr(34) & Format(V2.x, "000") & Chr(34) & " y=" & Chr(34) & Format(V2.y, "000") & Chr(34) & "/>")

            V1 = V2
        Next
        'PictureBox1.Image = bmp
        'graphics.Dispose()
        'PictureBox1.Refresh()
        'FileClose()
    End Sub

    Private Function GetAngle(ByVal V As vector) As Double
        '********************************************************************
        '********************************************************************

        Dim Vx As vector
        Vx.x = 1
        Vx.y = 0
        Vx.z = 0
        Vx.rho = 1
        Dim Vy As vector
        Vy.x = 0
        Vy.y = 1
        Vy.z = 0
        Vy.rho = 1
        Dim Vz As vector
        Vz.x = 0
        Vz.y = 0
        Vz.z = 1
        Vz.rho = 1

        'Dim V As vector

        'V.x = 0
        'V.y = -1
        'V.z = 0
        'V.rho = 1

        Dim scratchV As vector = Normalise(V)
        scratchV.rho = 1

        Dim scratchTheta As Double = Math.Acos(NewDotProduct(scratchV, Vx)) * 180 / Math.PI


        '*****************************************************
        '*****************************************************
        '                           
        '                            |
        '                            |
        '                    III     |       IV
        '                            |
        '                            |
        '                ------------------------- +x
        '                            |
        '                            |
        '                    II      |       I
        '                            |
        '                            |
        '                            |
        '                           +y
        '
        '*****************************************************
        '*****************************************************

        Select Case V.x
            Case Is > 0
                '   sector I or IV
                Select Case V.y
                    Case Is > 0
                        '   sector I = 0-90 degrees
                    Case Is < 0
                        '   sector II = 270-360 degrees
                        scratchTheta = -scratchTheta
                    Case Else
                        '   0 degrees
                End Select
            Case Is < 0
                '   sector II or IV
                Select Case V.y
                    Case Is > 0
                        '   sector IV = 180-270 degrees
                    Case Is < 0
                        '   sector II
                        scratchTheta = -scratchTheta
                    Case Else
                        '   180 degrees
                End Select
            Case Else
                '   lies on y axis
                Select Case V.y
                    Case Is > 0
                        '   90 degrees
                    Case Is < 0
                        '   270 degrees
                        scratchTheta = -scratchTheta
                    Case Else
                        '   impossible
                End Select
        End Select
        GetAngle = scratchTheta
    End Function

    Private Sub Cluster()
        Dim m As Integer
        Dim n As Integer
        Dim rho As Double
        Dim sp1, sp2 As Object
        Dim x1, y1, x2, y2 As Integer
        Dim xi, yi, xj, yj As Integer
        Dim BiggestRho As Double
        Dim BiggestRhoIndex(1) As Integer

        Dim SmallestiRho As Double
        'Dim SmallestiRhoIndex As Integer
        Dim SmallestjRho As Double
        'Dim SmallestjRhoIndex As Integer

        '   find the two most distant points as initial centroids for the two clusters

        BiggestRho = -1000000.0

        For m = 0 To ListBox2.Items.Count - 1

            sp1 = Split(ListBox2.Items.Item(m), Chr(9))
            x1 = CInt(sp1(0))
            y1 = CInt(sp1(1))

            For n = 0 To ListBox2.Items.Count - 1
                If m <> n Then
                    sp2 = Split(ListBox2.Items.Item(n), Chr(9))
                    x2 = CInt(sp2(0))
                    y2 = CInt(sp2(1))
                    rho = Distance(x1, y1, x2, y2)
                    If rho > BiggestRho Then
                        BiggestRho = rho
                        BiggestRhoIndex(0) = m
                        BiggestRhoIndex(1) = n
                    End If
                End If
            Next
        Next
        sp1 = Split(ListBox2.Items.Item(BiggestRhoIndex(0)), Chr(9))
        x1 = CInt(sp1(0))
        y1 = CInt(sp1(1))
        sp2 = Split(ListBox2.Items.Item(BiggestRhoIndex(1)), Chr(9))
        x2 = CInt(sp2(0))
        y2 = CInt(sp2(1))
        ListBox21.Items.Add("centroid m " & Chr(9) & Format(x1, "0000") & Chr(9) & Format(y1, "0000"))
        ListBox21.Items.Add("centroid n " & Chr(9) & Format(x2, "0000") & Chr(9) & Format(y2, "0000"))

        '   assign each point to one or the other centroid on the basis of distance



        For n = 0 To ListBox2.Items.Count - 1

            sp1 = Split(ListBox2.Items.Item(n), Chr(9))

            xi = CInt(sp1(0))
            yi = CInt(sp1(1))
            SmallestiRho = Distance(x1, y1, xi, yi)

            xj = CInt(sp1(0))
            yj = CInt(sp1(1))
            SmallestjRho = Distance(x2, y2, xj, yj)

            '   pick the smallest distance

            If SmallestiRho <= SmallestjRho Then
                ListBox2.Items.Item(n) = ListBox2.Items.Item(n) & Chr(9) & "1"
            Else
                ListBox2.Items.Item(n) = ListBox2.Items.Item(n) & Chr(9) & "2"
            End If
        Next

        '   now go through and throw out the 1's


        For n = ListBox2.Items.Count - 1 To 0 Step -1

            sp1 = Split(ListBox2.Items.Item(n), Chr(9))

            If sp1(2) = "1" Then
                ListBox2.Items.RemoveAt(n)
            End If
        Next

    End Sub



    Private Sub Button40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button40.Click
        '********************************************************************
        '********************************************************************
        '********************************************************************
        '********************************************************************


        Dim bmp As New Bitmap(PictureBox1.Image)
        Dim graphics As Graphics = graphics.FromImage(bmp)

        Dim CentreX As Integer = PictureBox1.Size.Width / 2
        Dim CentreY As Integer = PictureBox1.Size.Height / 2

        PenWidth = CSng(TextBox9.Text)

        Dim XDisplay1 As Single
        Dim YDisplay1 As Single

        Dim XDisplay2 As Single
        Dim YDisplay2 As Single

        Dim sp As Object

        Dim BoundariesFlag As Boolean = False

        Dim ii As Single
        Dim iii As Single
        'Dim j As Integer

        Dim x1, y1 As Integer ', x2, y2 As Integer

        'Dim rho11, rho12, rho21, rho22 As Double

        Dim LoopCount As Integer = 0
        Dim TraceCount1 As Integer = 0
        Dim TraceCount2 As Integer = 0

        'Dim rho As Double
        Dim EndsRho(3) As Double

        Dim RhoLimit As Double = 1 * 2 ^ 0.5

        '********************************************************************
        '********************************************************************

        graphics.Clear(Color.Ivory)

        GroupBox4.Visible = True
        GroupBox5.Visible = True
        GroupBox3.Visible = True
        GroupBox7.Visible = True

        Button8.Visible = True
        Button2.Visible = True

        Button4.Text = "Front"

        PinSpacing = 25.4 '* 10

        Dim xHoles As Integer = CInt(TextBox7.Text)
        Dim yHoles As Integer = CInt(TextBox8.Text)

        If xHoles < 25 Then xHoles = 25
        If yHoles < 5 Then yHoles = 5
        If xHoles > 85 Then xHoles = 85
        If yHoles > 60 Then yHoles = 60

        TextBox7.Text = Str(xHoles)
        TextBox8.Text = Str(yHoles)

        XBoardLimit = (xHoles + 1) * PinSpacing
        YBoardLimit = yHoles * PinSpacing

        TextBox26.Text = XBoardLimit
        TextBox27.Text = YBoardLimit

        TextBox26.Refresh()
        TextBox27.Refresh()

        OffsetX = 50 'CInt(TextBox28.Text)
        OffsetY = 50 'CInt(TextBox29.Text)
        DisplayScale = CSng(TextBox35.Text)



        PenWidth = CSng(TextBox9.Text)

        '********************************************************************
        '********************************************************************

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Clear()
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox19.Items.Clear()
        ListBox21.Items.Clear()


        Dim XOrigin, YOrigin As Double

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ComboBox7.Items.Clear()

        '********************************************************************
        '********************************************************************
        Dim InputPath As String = BasePath & "\Footprints\"
        ''Dim InputPath As String = "C:\Program Files\Clanking Replicators\Stripboard Designer\Footprints\"

        Dim InputString As String

        '   load in the footprints

        FileOpen(100, InputPath & "Footprints 01.txt", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Do Until EOF(100)
            InputString = LineInput(100)
            'ListBox19.Items.Add(InputString)
            If InStr(InputString, "<footprint = ") > 0 Then
                sp = Split(InputString, Chr(34))
                ComboBox7.Items.Add(sp(1))
            End If
        Loop
        ComboBox7.Text = "component footprints"
        FileClose(100)
        '********************************************************************
        '********************************************************************
        ''FileOpen(99, OutputPath & "output file0.txt", OpenMode.Output, OpenAccess.Write, OpenShare.Shared)
        ''FileOpen(100, OutputPath & "output file1.txt", OpenMode.Output, OpenAccess.Write, OpenShare.Shared)
        FileClose()
        'Dim OutputFileName As String = Mid(ComboBox1.Text, 1, ComboBox1.Text.Length - 4) & ".TXT"
        'FileOpen(101, OutputPath & "PCB.TXT", OpenMode.Output, OpenAccess.Write, OpenShare.Shared)



        ListBox20.Items.Clear()



        XOrigin = OffsetX
        YOrigin = OffsetY

        'x1 = Radius01 + OffsetX
        x1 = OffsetX
        y1 = YBoardLimit - OffsetY

        XDisplay1 = OffsetX / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY + PinSpacing) / DisplayScale
        XDisplay2 = OffsetX / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        XDisplay1 = OffsetX / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale
        XDisplay2 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        XDisplay1 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale
        XDisplay2 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY + PinSpacing) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        XDisplay1 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY + PinSpacing) / DisplayScale
        XDisplay2 = OffsetX / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY + PinSpacing) / DisplayScale

        ''Call NewDrawRedLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        'PictureBox1.Image = bmp
        'graphics.Dispose()
        'PictureBox1.Refresh()
        'Exit Sub

        PenWidth = 9

        Dim ii1 As Integer = 0.75 * PinSpacing
        Dim ii2 As Integer = XBoardLimit - 0.75 * PinSpacing

        ''For ii = PinSpacing To XBoardLimit - PinSpacing Step XBoardLimit - 2 * PinSpacing
        'For iii = PinSpacing To YBoardLimit - PinSpacing Step PinSpacing
        For iii = 1 To YBoardLimit Step PinSpacing
            XDisplay1 = (OffsetX + (ii1 - 5)) / DisplayScale
            YDisplay1 = (YBoardLimit + OffsetY - (iii)) / DisplayScale
            XDisplay2 = (OffsetX + (ii2 + 5)) / DisplayScale
            YDisplay2 = (YBoardLimit + OffsetY - (iii)) / DisplayScale

            Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "DarkGold", bmp)

            ListBox3.Items.Add("<strip>")
            ListBox3.Items.Add(Chr(9) & "<penwidth =" & Chr(34) & Str(PenWidth) & Chr(34) & "/>")
            ListBox3.Items.Add(Chr(9) & "<begin x=" & Chr(34) & Str(ii1) & Chr(34) & " y=" & Chr(34) & Str(iii) & Chr(34) & "/>")
            ListBox3.Items.Add(Chr(9) & "<end   x=" & Chr(34) & Str(ii2) & Chr(34) & " y=" & Chr(34) & Str(iii) & Chr(34) & "/>")
            ListBox3.Items.Add("</strip>")


            XDisplay1 = (OffsetX + (ii)) / DisplayScale
            YDisplay1 = (YBoardLimit + OffsetY - (iii - 5)) / DisplayScale
            XDisplay2 = (OffsetX + (ii)) / DisplayScale
            YDisplay2 = (YBoardLimit + OffsetY - (iii + 5)) / DisplayScale

            'Call NewDrawGoldLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)


            XDisplay1 = (ii) / DisplayScale + OffsetX
            YDisplay1 = YBoardLimit + OffsetY - (iii) / DisplayScale
            ListBox20.Items.Add(Str(ii) & Chr(9) & Str(iii) & Chr(9) & Str(XDisplay1) & Chr(9) & Str(YDisplay1))

        Next
        ''Next

        PenWidth = 5

        'For ii = PinSpacing To XBoardLimit - PinSpacing Step PinSpacing
        For ii = PinSpacing To XBoardLimit - PinSpacing / 2 Step PinSpacing
            'For iii = PinSpacing To YBoardLimit - PinSpacing Step PinSpacing
            For iii = 1 To YBoardLimit Step PinSpacing

                XDisplay1 = (OffsetX + (ii - 5)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - (iii)) / DisplayScale
                XDisplay2 = (OffsetX + (ii + 5)) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - (iii)) / DisplayScale

                'Call NewDrawGreenLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)

                XDisplay1 = (OffsetX + (ii)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - (iii)) / DisplayScale

                ''Call NewDrawBlackArcLineSegment(XDisplay1, YDisplay1, bmp)

                Call DrawCircle(XDisplay1, YDisplay1, "Black", bmp)

                XDisplay1 = (OffsetX + (ii)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - (iii - 5)) / DisplayScale
                XDisplay2 = (OffsetX + (ii)) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - (iii + 5)) / DisplayScale

                'Call NewDrawGreenLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)


                XDisplay1 = (ii) / DisplayScale + OffsetX
                YDisplay1 = YBoardLimit + OffsetY - (iii) / DisplayScale
                ListBox20.Items.Add(Str(ii) & Chr(9) & Str(iii) & Chr(9) & Str(XDisplay1) & Chr(9) & Str(YDisplay1))

            Next
        Next

        PenWidth = CSng(TextBox9.Text)

        PictureBox1.Image = bmp
        graphics.Dispose()
        PictureBox1.Refresh()

        Exit Sub

        '********************************************************************

    End Sub

    Private Sub Button41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ToolFlag = "break"
        PenWidth = CInt(TextBox24.Text)
        ListBox21.Items.Clear()
    End Sub

    Private Sub Button42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ToolFlag = "stripboard jumper"
        PenWidth = CInt(TextBox23.Text)
        ListBox21.Items.Clear()
    End Sub

    Private Sub PictureBox1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseDown
        Dim bmp As New Bitmap(PictureBox1.Image)
        Dim graphics As Graphics = graphics.FromImage(bmp)

        Dim ii, iii As Single
        Dim XDisplay1, YDisplay1 As Single
        Dim XDisplay2, YDisplay2 As Single

        Dim lastPenWidth As Single

        Dim sp As Object

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        'ListBox3.Items.Clear()

        x_md = e.X
        y_md = e.Y

        Dim x1, y1 As Double
        Dim x2, y2 As Double

        Dim x10, y10 As Double
        Dim x20, y20 As Double
        Dim x20last, y20last As Double
        Dim x20first, y20first As Double


        ' y = YBoardLimit + OffsetY  -YDisplay2*DisplayScale
        'XDisplay2*DisplayScale = (OffsetX + X) / 

        x10 = e.X * DisplayScale - OffsetX
        'XDisplay2*DisplayScale +OffsetX = XBoardLimit 

        y10 = YBoardLimit + OffsetY - e.Y * DisplayScale
        'y10 = (YBoardLimit + OffsetY - (e.Y + OffsetY)) * DisplayScale
        'x1 = e.X * DisplayScale - OffsetX
        'y1 = e.Y * DisplayScale

        TextBox14.Text = Str(x10)
        TextBox15.Text = Str(y10)

        'TextBox25.Text = Str(e.X) & " " & Str(e.Y)
        Try
            'TextBox25.Text = Str(PCBGrid(CInt(x10), CInt(y10)))
        Catch ex As Exception

        End Try


        'Exit Sub

        For ii = 0 To ListBox20.Items.Count - 1
            sp = Split(ListBox20.Items.Item(ii), Chr(9))
            Try
                x2 = CSng(sp(0))
                y2 = CSng(sp(1))
            Catch ex As Exception
                Exit Sub
            End Try

            'ListBox1.Items.Add(Format(Distance(x1, y1, x2, y2), "0000.0") & Chr(9) & Str(ii) & Chr(9) & Str(iii) & _
            '            Chr(9) & Format(x2, "0000.0") & Chr(9) & Format(y2, "0000.0") & _
            '            Chr(9) & Format(x1, "0000.0") & Chr(9) & Format(y1, "0000.0"))
            If Distance(x10, y10, x2, y2) <= PinSpacing / 2 Then
                XDisplay1 = (OffsetX + x2) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y2) / DisplayScale
                ListBox1.Items.Add(Format(Distance(x10, y10, x2, y2), "0000.0") & Chr(9) & sp(0) & Chr(9) & sp(1) & _
                                        Chr(9) & Format(x2, "0000.0") & Chr(9) & Format(y2, "0000.0") & _
                                        Chr(9) & Format(x10, "0000.0") & Chr(9) & Format(y10, "0000.0"))
                'Call NewDrawRedArcLineSegment(XDisplay1, YDisplay1, bmp)

                If ToolFlag = "footprint" Then
                    '********************************************************************
                    '********************************************************************
                    ''Dim InputPath As String = Mid(CurDir, 1, Len(CurDir) - 8) & "Stripboards\Footprints\"
                    Dim InputPath As String = BasePath & "\Footprints\"

                    Dim InputString As String
                    Dim FootprintOnBoard As Boolean = True

                    '   check to see if the footprint is actually completely on the board

                    FileOpen(100, InputPath & "Footprints 01.txt", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

                    Do Until EOF(100)
                        InputString = LineInput(100)
                        ListBox19.Items.Add(InputString)
                        If InStr(InputString, "<footprint = ") > 0 And InStr(InputString, ComboBox7.Text) > 0 Then
                            ListBox19.Items.Add(InputString)
                            Do Until InStr(InputString, "</footprint>") > 0
                                InputString = LineInput(100)
                                ListBox19.Items.Add(InputString)
                                If InStr(InputString, "<perimeter>") > 0 Then
                                    Dim PerimeterBegin As Boolean = False
                                    Do Until InStr(InputString, "</perimeter>") > 0
                                        InputString = LineInput(100)
                                        If InStr(InputString, "</perimeter>") = 0 Then
                                            ListBox2.Items.Add(InputString)
                                            sp = Split(InputString, Chr(34))
                                            x20first = (x2 + CSng(sp(1)) * PinSpacing) / DisplayScale
                                            y20first = (y2 - CSng(sp(3)) * PinSpacing) / DisplayScale
                                            x20last = (x2 + CSng(sp(5)) * PinSpacing) / DisplayScale
                                            y20last = (y2 - CSng(sp(7)) * PinSpacing) / DisplayScale
                                            'Call NewDrawBlueLineSegment(x20first, y20first, x20last, y20last, bmp)
                                        End If
                                    Loop
                                ElseIf InStr(InputString, "<pins>") > 0 Then
                                    'GoTo skip01
                                    Do Until InStr(InputString, "</pins>") > 0
                                        InputString = LineInput(100)
                                        If InStr(InputString, "</pins>") = 0 Then
                                            sp = Split(InputString, Chr(34))
                                            'ListBox3.Items.Add(InputString)
                                            x20 = x2 + CSng(sp(3)) * PinSpacing '/ DisplayScale
                                            y20 = y2 + CSng(sp(5)) * PinSpacing '/ DisplayScale

                                            XDisplay1 = (OffsetX + x20) / DisplayScale
                                            YDisplay1 = (YBoardLimit - OffsetY - y20) / DisplayScale

                                            XDisplay2 = (OffsetX + (XBoardLimit - PinSpacing)) / DisplayScale
                                            YDisplay2 = (YBoardLimit - OffsetY - YBoardLimit) / DisplayScale

                                            If XDisplay1 > XDisplay2 Or YDisplay1 < YDisplay2 Then
                                                FootprintOnBoard = False
                                                FileClose(100)
                                                GoTo abort01
                                            ElseIf XDisplay1 < (OffsetX) Or YDisplay1 > YBoardLimit + (OffsetY + 0.001) Then
                                                FootprintOnBoard = False
                                                FileClose(100)
                                                GoTo abort01
                                            Else
                                                'Call NewDrawRedArcLineSegment(x20, y20, bmp)
                                            End If

                                        End If
                                    Loop
skip01:
                                End If

                            Loop

                        End If
                    Loop
                    FileClose(100)
                    '********************************************************************
                    '********************************************************************

                    '   load in the footprints

                    InputPath = BasePath & "\Footprints\"


                    FileOpen(100, InputPath & "Footprints 01.txt", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

                    Do Until EOF(100)
                        InputString = LineInput(100)
                        ListBox19.Items.Add(InputString)
                        If InStr(InputString, "<footprint = ") > 0 And InStr(InputString, ComboBox7.Text) > 0 Then
                            ListBox3.Items.Add(InputString)
                            ListBox3.Items.Add(Chr(9) & "<location x=" & Chr(34) & Str(x2) & Chr(34) & " y=" & Chr(34) & Str(y2) & Chr(34) & "/>")

                            ListBox19.Items.Add(InputString)
                            Do Until InStr(InputString, "</footprint>") > 0
                                InputString = LineInput(100)
                                ListBox19.Items.Add(InputString)
                                If InStr(InputString, "<perimeter>") > 0 Then
                                    Dim PerimeterBegin As Boolean = False
                                    Do Until InStr(InputString, "</perimeter>") > 0
                                        InputString = LineInput(100)
                                        If InStr(InputString, "</perimeter>") = 0 Then
                                            ListBox2.Items.Add(InputString)
                                            sp = Split(InputString, Chr(34))
                                            x20first = x2 + CSng(sp(1)) * PinSpacing ' / DisplayScale
                                            y20first = y2 + CSng(sp(3)) * PinSpacing ' / DisplayScale
                                            x20last = x2 + CSng(sp(5)) * PinSpacing ' / DisplayScale
                                            y20last = y2 + CSng(sp(7)) * PinSpacing ' / DisplayScale

                                            XDisplay1 = (OffsetX + x20first) / DisplayScale
                                            YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                                            XDisplay2 = (OffsetX + x20last) / DisplayScale
                                            YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale
                                            PenWidth = 3
                                            ''Call NewDrawBlueLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
                                            Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Blue", bmp)
                                            PenWidth = CSng(TextBox9.Text)
                                        End If
                                    Loop
                                ElseIf InStr(InputString, "<pins>") > 0 Then
                                    Do Until InStr(InputString, "</pins>") > 0
                                        InputString = LineInput(100)
                                        If InStr(InputString, "</pins>") = 0 Then
                                            sp = Split(InputString, Chr(34))
                                            'ListBox3.Items.Add(InputString)
                                            x20 = x2 + CSng(sp(3)) * PinSpacing '/ DisplayScale
                                            y20 = y2 + CSng(sp(5)) * PinSpacing '/ DisplayScale

                                            x20first = x20 - 10 '/ 2
                                            y20first = y20
                                            'sp = Split(ListBox21.Items.Item(iii), Chr(9))
                                            x20last = x20 + 10 '/ 2
                                            y20last = y20
                                            lastPenWidth = PenWidth
                                            PenWidth = 10 'CSng(TextBox24.Text)

                                            XDisplay1 = (OffsetX + x20first) / DisplayScale
                                            YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                                            XDisplay2 = (OffsetX + x20last) / DisplayScale
                                            YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                                            'Call NewDrawGoldLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
                                            Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Gold", bmp)
                                            PenWidth = lastPenWidth

                                            ListBox3.Items.Add(Chr(9) & "<pin pad x=" & Chr(34) & Str(x20) & Chr(34) & " y=" & Chr(34) & Str(y20) & Chr(34) & "/>")

                                            XDisplay1 = (OffsetX + x20) / DisplayScale
                                            YDisplay1 = (YBoardLimit + OffsetY - y20) / DisplayScale
                                            ''Call NewDrawRedArcLineSegment(XDisplay1, YDisplay1, bmp)

                                            Call DrawCircle(XDisplay1, YDisplay1, "Red", bmp)
                                        End If
                                    Loop
                                    ListBox3.Items.Add("</footprint>")
                                End If

                            Loop
                        End If
                    Loop
                    FileClose(100)
abort01:
                    '********************************************************************
                    '********************************************************************
                ElseIf ToolFlag = "trace" Then
                    '********************************************************************
                    '********************************************************************
                    ListBox21.Items.Add(Str(x2) & Chr(9) & Str(y2))

                    If ListBox21.Items.Count > 1 Then
                        For iii = ListBox21.Items.Count - 1 To ListBox21.Items.Count - 1
                            sp = Split(ListBox21.Items.Item(iii - 1), Chr(9))
                            x20first = CSng(sp(0))
                            y20first = CSng(sp(1))
                            sp = Split(ListBox21.Items.Item(iii), Chr(9))
                            x20last = CSng(sp(0))
                            y20last = CSng(sp(1))

                            ListBox3.Items.Add("<trace>")
                            ListBox3.Items.Add(Chr(9) & "<penwidth =" & Chr(34) & Str(PenWidth) & Chr(34) & "/>")
                            ListBox3.Items.Add(Chr(9) & "<begin x=" & Chr(34) & Str(x20first) & Chr(34) & " y=" & Chr(34) & Str(y20first) & Chr(34) & "/>")
                            ListBox3.Items.Add(Chr(9) & "<end   x=" & Chr(34) & Str(x20last) & Chr(34) & " y=" & Chr(34) & Str(y20last) & Chr(34) & "/>")
                            ListBox3.Items.Add("</trace>")

                            XDisplay1 = (OffsetX + x20first) / DisplayScale
                            YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                            XDisplay2 = (OffsetX + x20last) / DisplayScale
                            YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                            'Call NewDrawGoldLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
                            'Call NewDrawGoldArcLineSegment(XDisplay1, YDisplay1, bmp)
                            'Call NewDrawGoldArcLineSegment(XDisplay2, YDisplay2, bmp)

                            Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "Gold", bmp)

                            Call DrawCircle(XDisplay1, YDisplay1, "Gold", bmp)
                            Call DrawCircle(XDisplay2, YDisplay2, "Gold", bmp)

                        Next
                    End If
                    '********************************************************************
                    '********************************************************************

                ElseIf ToolFlag = "jumper" Then
                    '********************************************************************
                    '********************************************************************
                    ListBox21.Items.Add(Str(x2) & Chr(9) & Str(y2))

                    If ListBox21.Items.Count > 1 Then
                        For iii = ListBox21.Items.Count - 1 To ListBox21.Items.Count - 1
                            sp = Split(ListBox21.Items.Item(iii - 1), Chr(9))
                            x20first = CSng(sp(0))
                            y20first = CSng(sp(1))
                            sp = Split(ListBox21.Items.Item(iii), Chr(9))
                            x20last = CSng(sp(0))
                            y20last = CSng(sp(1))


                            ListBox3.Items.Add("<jumper>")
                            ListBox3.Items.Add(Chr(9) & "<penwidth =" & Chr(34) & Str(PenWidth) & Chr(34) & "/>")
                            ListBox3.Items.Add(Chr(9) & "<begin x=" & Chr(34) & Str(x20first) & Chr(34) & " y=" & Chr(34) & Str(y20first) & Chr(34) & "/>")
                            ListBox3.Items.Add(Chr(9) & "<end   x=" & Chr(34) & Str(x20last) & Chr(34) & " y=" & Chr(34) & Str(y20last) & Chr(34) & "/>")
                            ListBox3.Items.Add("</jumper>")

                            XDisplay1 = (OffsetX + x20first) / DisplayScale
                            YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                            XDisplay2 = (OffsetX + x20last) / DisplayScale
                            YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                            Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Green", bmp)

                        Next
                    End If
                    '********************************************************************
                    '********************************************************************
                ElseIf ToolFlag = "stripboard jumper" Then
                    '********************************************************************
                    '********************************************************************
                    ListBox21.Items.Add(Str(x2) & Chr(9) & Str(y2))

                    If ListBox21.Items.Count > 1 Then
                        For iii = ListBox21.Items.Count - 1 To ListBox21.Items.Count - 1
                            sp = Split(ListBox21.Items.Item(iii - 1), Chr(9))
                            x20first = CSng(sp(0))
                            y20first = CSng(sp(1))
                            sp = Split(ListBox21.Items.Item(iii), Chr(9))
                            x20last = CSng(sp(0))
                            y20last = CSng(sp(1))


                            ListBox3.Items.Add("<stripboard jumper>")
                            ListBox3.Items.Add(Chr(9) & "<penwidth =" & Chr(34) & Str(PenWidth) & Chr(34) & "/>")
                            ListBox3.Items.Add(Chr(9) & "<begin x=" & Chr(34) & Str(x20first) & Chr(34) & " y=" & Chr(34) & Str(y20first) & Chr(34) & "/>")
                            ListBox3.Items.Add(Chr(9) & "<end   x=" & Chr(34) & Str(x20first) & Chr(34) & " y=" & Chr(34) & Str(y20last) & Chr(34) & "/>")
                            ListBox3.Items.Add("</stripboard jumper>")

                            XDisplay1 = (OffsetX + x20first) / DisplayScale
                            YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                            XDisplay2 = (OffsetX + x20last) / DisplayScale
                            YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                            'Call NewDrawOrangeLineSegment(XDisplay1, YDisplay1, XDisplay1, YDisplay2, bmp)
                            Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "HotPink", bmp)

                            Call DrawCircle(XDisplay1, YDisplay1, "HotPink", bmp)
                            Call DrawCircle(XDisplay1, YDisplay2, "HotPink", bmp)

                        Next
                    End If
                    'ToolFlag = ""
                    '********************************************************************
                    '********************************************************************
                ElseIf ToolFlag = "stripboard jumper - 0 - +5 v" Then
                    '********************************************************************
                    '********************************************************************
                    ListBox21.Items.Add(Str(x2) & Chr(9) & Str(y2))

                    If ListBox21.Items.Count > 1 Then
                        For iii = ListBox21.Items.Count - 1 To ListBox21.Items.Count - 1
                            sp = Split(ListBox21.Items.Item(iii - 1), Chr(9))
                            x20first = CSng(sp(0))
                            y20first = CSng(sp(1))
                            sp = Split(ListBox21.Items.Item(iii), Chr(9))
                            x20last = CSng(sp(0))
                            y20last = CSng(sp(1))


                            ListBox3.Items.Add("<stripboard jumper - 0 - +5 v>")
                            ListBox3.Items.Add(Chr(9) & "<penwidth =" & Chr(34) & Str(PenWidth) & Chr(34) & "/>")
                            ListBox3.Items.Add(Chr(9) & "<begin x=" & Chr(34) & Str(x20first) & Chr(34) & " y=" & Chr(34) & Str(y20first) & Chr(34) & "/>")
                            ListBox3.Items.Add(Chr(9) & "<end   x=" & Chr(34) & Str(x20first) & Chr(34) & " y=" & Chr(34) & Str(y20last) & Chr(34) & "/>")
                            ListBox3.Items.Add("</stripboard jumper - 0 - +5 v>")

                            XDisplay1 = (OffsetX + x20first) / DisplayScale
                            YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                            XDisplay2 = (OffsetX + x20last) / DisplayScale
                            YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                            'Call NewDrawBlackLineSegment(XDisplay1, YDisplay1, XDisplay1, YDisplay2, bmp)
                            Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "Black", bmp)

                            Call DrawCircle(XDisplay1, YDisplay1, "Black", bmp)
                            Call DrawCircle(XDisplay1, YDisplay2, "Black", bmp)
                        Next
                    End If
                    'ToolFlag = ""
                    '********************************************************************
                    '********************************************************************
                ElseIf ToolFlag = "stripboard jumper - 0 - +12 v" Then
                    '********************************************************************
                    '********************************************************************
                    ListBox21.Items.Add(Str(x2) & Chr(9) & Str(y2))

                    If ListBox21.Items.Count > 1 Then
                        For iii = ListBox21.Items.Count - 1 To ListBox21.Items.Count - 1
                            sp = Split(ListBox21.Items.Item(iii - 1), Chr(9))
                            x20first = CSng(sp(0))
                            y20first = CSng(sp(1))
                            sp = Split(ListBox21.Items.Item(iii), Chr(9))
                            x20last = CSng(sp(0))
                            y20last = CSng(sp(1))


                            ListBox3.Items.Add("<stripboard jumper - 0 - +12 v>")
                            ListBox3.Items.Add(Chr(9) & "<penwidth =" & Chr(34) & Str(PenWidth) & Chr(34) & "/>")
                            ListBox3.Items.Add(Chr(9) & "<begin x=" & Chr(34) & Str(x20first) & Chr(34) & " y=" & Chr(34) & Str(y20first) & Chr(34) & "/>")
                            ListBox3.Items.Add(Chr(9) & "<end   x=" & Chr(34) & Str(x20first) & Chr(34) & " y=" & Chr(34) & Str(y20last) & Chr(34) & "/>")
                            ListBox3.Items.Add("</stripboard jumper - 0 - +12 v>")

                            XDisplay1 = (OffsetX + x20first) / DisplayScale
                            YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                            XDisplay2 = (OffsetX + x20last) / DisplayScale
                            YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                            Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "Purple", bmp)

                            Call DrawCircle(XDisplay1, YDisplay1, "Purple", bmp)
                            Call DrawCircle(XDisplay1, YDisplay2, "Purple", bmp)

                        Next
                    End If
                    'ToolFlag = ""
                    '********************************************************************
                    '********************************************************************
                ElseIf ToolFlag = "stripboard jumper - +5 v" Then
                    '********************************************************************
                    '********************************************************************
                    ListBox21.Items.Add(Str(x2) & Chr(9) & Str(y2))

                    If ListBox21.Items.Count > 1 Then
                        For iii = ListBox21.Items.Count - 1 To ListBox21.Items.Count - 1
                            sp = Split(ListBox21.Items.Item(iii - 1), Chr(9))
                            x20first = CSng(sp(0))
                            y20first = CSng(sp(1))
                            sp = Split(ListBox21.Items.Item(iii), Chr(9))
                            x20last = CSng(sp(0))
                            y20last = CSng(sp(1))


                            ListBox3.Items.Add("<stripboard jumper - +5 v>")
                            ListBox3.Items.Add(Chr(9) & "<penwidth =" & Chr(34) & Str(PenWidth) & Chr(34) & "/>")
                            ListBox3.Items.Add(Chr(9) & "<begin x=" & Chr(34) & Str(x20first) & Chr(34) & " y=" & Chr(34) & Str(y20first) & Chr(34) & "/>")
                            ListBox3.Items.Add(Chr(9) & "<end   x=" & Chr(34) & Str(x20first) & Chr(34) & " y=" & Chr(34) & Str(y20last) & Chr(34) & "/>")
                            ListBox3.Items.Add("</stripboard jumper - +5 v>")

                            XDisplay1 = (OffsetX + x20first) / DisplayScale
                            YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                            XDisplay2 = (OffsetX + x20last) / DisplayScale
                            YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                            Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "DarkRed", bmp)

                            Call DrawCircle(XDisplay1, YDisplay1, "DarkRed", bmp)
                            Call DrawCircle(XDisplay1, YDisplay2, "DarkRed", bmp)

                        Next
                    End If
                    'ToolFlag = ""
                    '********************************************************************
                    '********************************************************************
                ElseIf ToolFlag = "stripboard jumper - +12 v" Then
                    '********************************************************************
                    '********************************************************************
                    ListBox21.Items.Add(Str(x2) & Chr(9) & Str(y2))

                    If ListBox21.Items.Count > 1 Then
                        For iii = ListBox21.Items.Count - 1 To ListBox21.Items.Count - 1
                            sp = Split(ListBox21.Items.Item(iii - 1), Chr(9))
                            x20first = CSng(sp(0))
                            y20first = CSng(sp(1))
                            sp = Split(ListBox21.Items.Item(iii), Chr(9))
                            x20last = CSng(sp(0))
                            y20last = CSng(sp(1))


                            ListBox3.Items.Add("<stripboard jumper - +12 v>")
                            ListBox3.Items.Add(Chr(9) & "<penwidth =" & Chr(34) & Str(PenWidth) & Chr(34) & "/>")
                            ListBox3.Items.Add(Chr(9) & "<begin x=" & Chr(34) & Str(x20first) & Chr(34) & " y=" & Chr(34) & Str(y20first) & Chr(34) & "/>")
                            ListBox3.Items.Add(Chr(9) & "<end   x=" & Chr(34) & Str(x20first) & Chr(34) & " y=" & Chr(34) & Str(y20last) & Chr(34) & "/>")
                            ListBox3.Items.Add("</stripboard jumper - +12 v>")

                            XDisplay1 = (OffsetX + x20first) / DisplayScale
                            YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                            XDisplay2 = (OffsetX + x20last) / DisplayScale
                            YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                            Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "Red", bmp)

                            Call DrawCircle(XDisplay1, YDisplay1, "Red", bmp)
                            Call DrawCircle(XDisplay1, YDisplay2, "Red", bmp)

                        Next
                    End If
                    'ToolFlag = ""
                    '********************************************************************
                    '********************************************************************
                ElseIf ToolFlag = "stripboard jumper - gnd" Then
                    '********************************************************************
                    '********************************************************************
                    ListBox21.Items.Add(Str(x2) & Chr(9) & Str(y2))

                    If ListBox21.Items.Count > 1 Then
                        For iii = ListBox21.Items.Count - 1 To ListBox21.Items.Count - 1
                            sp = Split(ListBox21.Items.Item(iii - 1), Chr(9))
                            x20first = CSng(sp(0))
                            y20first = CSng(sp(1))
                            sp = Split(ListBox21.Items.Item(iii), Chr(9))
                            x20last = CSng(sp(0))
                            y20last = CSng(sp(1))


                            ListBox3.Items.Add("<stripboard jumper - gnd>")
                            ListBox3.Items.Add(Chr(9) & "<penwidth =" & Chr(34) & Str(PenWidth) & Chr(34) & "/>")
                            ListBox3.Items.Add(Chr(9) & "<begin x=" & Chr(34) & Str(x20first) & Chr(34) & " y=" & Chr(34) & Str(y20first) & Chr(34) & "/>")
                            ListBox3.Items.Add(Chr(9) & "<end   x=" & Chr(34) & Str(x20first) & Chr(34) & " y=" & Chr(34) & Str(y20last) & Chr(34) & "/>")
                            ListBox3.Items.Add("</stripboard jumper- gnd>")

                            XDisplay1 = (OffsetX + x20first) / DisplayScale
                            YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                            XDisplay2 = (OffsetX + x20last) / DisplayScale
                            YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                            Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "DarkGreen", bmp)

                            Call DrawCircle(XDisplay1, YDisplay1, "DarkGreen", bmp)
                            Call DrawCircle(XDisplay1, YDisplay2, "DarkGreen", bmp)

                            ToolFlag = ""

                        Next
                    End If
                    'ToolFlag = ""
                    '********************************************************************
                    '********************************************************************
                ElseIf ToolFlag = "pad" Then
                    '********************************************************************
                    '********************************************************************
                    'ListBox21.Items.Add(Str(x2) & Chr(9) & Str(y2))

                    'If ListBox21.Items.Count > 1 Then
                    'For iii = 1 To ListBox21.Items.Count - 1
                    'sp = Split(ListBox21.Items.Item(iii - 1), Chr(9))
                    x20first = x2 - CSng(TextBox24.Text) '/ 2
                    y20first = y2
                    'sp = Split(ListBox21.Items.Item(iii), Chr(9))
                    x20last = x2 + CSng(TextBox24.Text) '/ 2
                    y20last = y2
                    lastPenWidth = PenWidth
                    PenWidth = CSng(TextBox24.Text)

                    ListBox3.Items.Add("<pad>")
                    ListBox3.Items.Add(Chr(9) & "<penwidth =" & Chr(34) & Str(PenWidth) & Chr(34) & "/>")
                    ListBox3.Items.Add(Chr(9) & "<begin x=" & Chr(34) & Str(x20first) & Chr(34) & " y=" & Chr(34) & Str(y20first) & Chr(34) & "/>")
                    ListBox3.Items.Add(Chr(9) & "<end   x=" & Chr(34) & Str(x20last) & Chr(34) & " y=" & Chr(34) & Str(y20last) & Chr(34) & "/>")
                    ListBox3.Items.Add("</pad>")

                    XDisplay1 = (OffsetX + x20first) / DisplayScale
                    YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                    XDisplay2 = (OffsetX + x20last) / DisplayScale
                    YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                    'Call NewDrawGoldLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
                    Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Gold", bmp)
                    PenWidth = lastPenWidth
                    'Next
                    'End If
                    '********************************************************************
                    '********************************************************************
                ElseIf ToolFlag = "break" Then
                    '********************************************************************
                    '********************************************************************
                    'ListBox21.Items.Add(Str(x2) & Chr(9) & Str(y2))

                    'If ListBox21.Items.Count > 1 Then
                    'For iii = 1 To ListBox21.Items.Count - 1
                    'sp = Split(ListBox21.Items.Item(iii - 1), Chr(9))
                    x20first = x2 - CSng(TextBox24.Text) '/ 2
                    y20first = y2
                    'sp = Split(ListBox21.Items.Item(iii), Chr(9))
                    x20last = x2 + CSng(TextBox24.Text) '/ 2
                    y20last = y2
                    lastPenWidth = PenWidth
                    PenWidth = 9

                    ListBox3.Items.Add("<break>")
                    ListBox3.Items.Add(Chr(9) & "<penwidth =" & Chr(34) & Str(PenWidth) & Chr(34) & "/>")
                    ListBox3.Items.Add(Chr(9) & "<begin x=" & Chr(34) & Str(x20first) & Chr(34) & " y=" & Chr(34) & Str(y20first) & Chr(34) & "/>")
                    ListBox3.Items.Add(Chr(9) & "<end   x=" & Chr(34) & Str(x20last) & Chr(34) & " y=" & Chr(34) & Str(y20last) & Chr(34) & "/>")
                    ListBox3.Items.Add("</break>")

                    XDisplay1 = (OffsetX + x20first) / DisplayScale
                    YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                    XDisplay2 = (OffsetX + x20last) / DisplayScale
                    YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                    'Call NewDrawIvoryLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
                    Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Ivory", bmp)
                    PenWidth = CSng(TextBox24.Text)
                    'Next
                    'End If
                    '********************************************************************
                    '********************************************************************
                ElseIf ToolFlag = "twixt break" Then
                    '********************************************************************
                    '********************************************************************
                    'ListBox21.Items.Add(Str(x2) & Chr(9) & Str(y2))

                    'If ListBox21.Items.Count > 1 Then
                    'For iii = 1 To ListBox21.Items.Count - 1
                    'sp = Split(ListBox21.Items.Item(iii - 1), Chr(9))
                    x20first = x2 + 25.4 / 2
                    y20first = y2 - 25.4 / 2
                    'sp = Split(ListBox21.Items.Item(iii), Chr(9))
                    x20last = x2 + 25.4 / 2
                    y20last = y2 + 25.4 / 2
                    lastPenWidth = PenWidth
                    PenWidth = 2

                    ListBox3.Items.Add("<twixt break>")
                    ListBox3.Items.Add(Chr(9) & "<penwidth =" & Chr(34) & Str(PenWidth) & Chr(34) & "/>")
                    ListBox3.Items.Add(Chr(9) & "<begin x=" & Chr(34) & Str(x20first) & Chr(34) & " y=" & Chr(34) & Str(y20first) & Chr(34) & "/>")
                    ListBox3.Items.Add(Chr(9) & "<end   x=" & Chr(34) & Str(x20last) & Chr(34) & " y=" & Chr(34) & Str(y20last) & Chr(34) & "/>")
                    ListBox3.Items.Add("</twixt break>")

                    XDisplay1 = (OffsetX + x20first) / DisplayScale
                    YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                    XDisplay2 = (OffsetX + x20last) / DisplayScale
                    YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                    'Call NewDrawIvoryLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
                    Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Ivory", bmp)
                    PenWidth = CSng(TextBox24.Text)
                    'Next
                    'End If
                    '********************************************************************
                    '********************************************************************
                ElseIf ToolFlag = "select" Then
                    '********************************************************************
                    '********************************************************************
                    'Dim sp As Object
                    Dim jj As Integer
                    'Dim x1, y1 As Single
                    'Dim x2, y2 As Single
                    Dim x3, y3 As Single
                    '********************************************************************
                    '********************************************************************

                    '********************************************************************
                    '********************************************************************

                    x1 = CSng(TextBox14.Text)
                    y1 = CSng(TextBox15.Text)

                    'REDRAW THE DESIGN

                    For jj = 0 To ListBox3.Items.Count - 1
                        If InStr(ListBox3.Items.Item(jj), "<footprint = ") > 0 Then

                            '   get the location

                            sp = Split(ListBox3.Items.Item(jj + 1), Chr(34))

                            x2 = CSng(sp(1))
                            y2 = CSng(sp(3))
                            If Distance(x1, y1, x2, y2) <= PinSpacing / 2 Then
                                'ListBox3.Items.RemoveAt(jj + 2)
                                'ListBox3.Items.RemoveAt(jj + 1)
DoTheNextOne:
                                If InStr(ListBox3.Items.Item(jj), "</footprint>") = 0 Then
                                    ListBox3.Items.RemoveAt(jj)
                                    GoTo DoTheNextOne
                                Else
                                    ListBox3.Items.RemoveAt(jj)
                                    Exit For
                                End If
                            End If
                        ElseIf InStr(ListBox3.Items.Item(jj), "<stripboard jumper>") > 0 Then

                            '   get the location

                            sp = Split(ListBox3.Items.Item(jj + 2), Chr(34))

                            x2 = CSng(sp(1))
                            y2 = CSng(sp(3))
                            sp = Split(ListBox3.Items.Item(jj + 3), Chr(34))

                            x3 = CSng(sp(1))
                            y3 = CSng(sp(3))
                            If Distance(x1, y1, x2, y2) + Distance(x1, y1, x3, y3) <= _
                            1.02 * (Distance(x2, y2, x3, y3)) Then
                                ListBox3.Items.RemoveAt(jj + 4)
                                ListBox3.Items.RemoveAt(jj + 3)
                                ListBox3.Items.RemoveAt(jj + 2)
                                ListBox3.Items.RemoveAt(jj + 1)
                                ListBox3.Items.RemoveAt(jj)
                                Exit For
                            End If
                        ElseIf InStr(ListBox3.Items.Item(jj), "<stripboard jumper -") > 0 Then

                            '   get the location

                            sp = Split(ListBox3.Items.Item(jj + 2), Chr(34))

                            x2 = CSng(sp(1))
                            y2 = CSng(sp(3))
                            sp = Split(ListBox3.Items.Item(jj + 3), Chr(34))

                            x3 = CSng(sp(1))
                            y3 = CSng(sp(3))
                            If Distance(x1, y1, x2, y2) + Distance(x1, y1, x3, y3) <= _
                            1.02 * (Distance(x2, y2, x3, y3)) Then
                                ListBox3.Items.RemoveAt(jj + 4)
                                ListBox3.Items.RemoveAt(jj + 3)
                                ListBox3.Items.RemoveAt(jj + 2)
                                ListBox3.Items.RemoveAt(jj + 1)
                                ListBox3.Items.RemoveAt(jj)
                                Exit For
                            End If
                        ElseIf InStr(ListBox3.Items.Item(jj), "<break>") > 0 Then

                            '   get the location

                            sp = Split(ListBox3.Items.Item(jj + 2), Chr(34))

                            x2 = CSng(sp(1))
                            y2 = CSng(sp(3))
                            sp = Split(ListBox3.Items.Item(jj + 3), Chr(34))

                            x3 = CSng(sp(1))
                            y3 = CSng(sp(3))
                            If Distance(x1, y1, x2, y2) + Distance(x1, y1, x3, y3) <= _
                            1.05 * (Distance(x2, y2, x3, y3)) Then
                                ListBox3.Items.RemoveAt(jj + 4)
                                ListBox3.Items.RemoveAt(jj + 3)
                                ListBox3.Items.RemoveAt(jj + 2)
                                ListBox3.Items.RemoveAt(jj + 1)
                                ListBox3.Items.RemoveAt(jj)
                                Exit For
                            End If
                        End If
                    Next jj

                    Call RefreshStripboardDisplay(0)
                    Exit Sub

                End If
                PictureBox1.Image = bmp
                graphics.Dispose()
                PictureBox1.Refresh()
                Exit Sub
            End If
        Next

    End Sub


    Private Sub Button43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim bmp As New Bitmap(PictureBox1.Image)
        Dim graphics As Graphics = graphics.FromImage(bmp)

        'graphics.Clear(Color.Black)

        PictureBox1.Image = bmp
        graphics.Dispose()
        PictureBox1.Refresh()

        '********************************************************************

        'PictureBox1.Width = 1600
        'PictureBox1.Height = 1600

        '********************************************************************
        ''Dim InputPath As String = Mid(CurDir, 1, Len(CurDir) - 8) & "Stripboards\Boards\"

        Dim InputPath As String = BasePath & "\Footprints\"

        Dim InputFileName As String = OpenFileDialog1.FileName
        Dim InputString As String


        PenWidth = CSng(TextBox9.Text)


        Dim XDisplay1 As Single
        Dim YDisplay1 As Single

        Dim XDisplay2 As Single
        Dim YDisplay2 As Single
        Dim sp As Object

        Dim BoundariesFlag As Boolean = False

        Dim ii As Single
        Dim iii As Single
        'Dim j As Integer

        Dim x1, y1 As Integer

        Dim LoopCount As Integer = 0
        Dim TraceCount1 As Integer = 0
        Dim TraceCount2 As Integer = 0

        Dim EndsRho(3) As Double
        Dim lastPenWidth As Object

        Dim RhoLimit As Double = 1 * 2 ^ 0.5

        '********************************************************************
        '********************************************************************
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        'ListBox3.Items.Clear()
        ListBox4.Items.Clear()
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox19.Items.Clear()
        ListBox21.Items.Clear()

        Dim XOrigin, YOrigin As Double

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ComboBox7.Items.Clear()
        '********************************************************************
        '********************************************************************

        '   load in the footprints

        '   get the display scale

        InputPath = BasePath & "\Footprints\"


        FileOpen(100, InputPath & InputFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Do Until EOF(100)
            InputString = LineInput(100)
            If InStr(InputString, "<DisplayScale = ") > 0 Then
                sp = Split(InputString, Chr(34))
                DisplayScale = CSng(TextBox35.Text) 'CSng(sp(1))
                'TextBox30.Text = 1 / DisplayScale
                Exit Do
            End If
        Loop
        FileClose(100)

        '   get the pin spacing

        FileOpen(100, InputPath & InputFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Do Until EOF(100)
            InputString = LineInput(100)
            If InStr(InputString, "<PinSpacing") > 0 Then
                sp = Split(InputString, Chr(34))
                PinSpacing = CSng(sp(1))
                Exit Do
            End If
        Loop
        FileClose(100)

        '   get the board dimensions

        FileOpen(100, InputPath & InputFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Do Until EOF(100)
            InputString = LineInput(100)
            If InStr(InputString, "<Board dimensions") > 0 Then
                sp = Split(InputString, Chr(34))
                XBoardLimit = CSng(sp(1))
                YBoardLimit = CSng(sp(3))
                TextBox26.Text = XBoardLimit
                TextBox27.Text = YBoardLimit
                Exit Do
            End If
        Loop
        FileClose(100)

        '   get the offsets

        FileOpen(100, InputPath & InputFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Do Until EOF(100)
            InputString = LineInput(100)
            If InStr(InputString, "<Display offsets") > 0 Then
                sp = Split(InputString, Chr(34))
                OffsetX = CSng(sp(1))
                OffsetY = CSng(sp(3))
                'TextBox28.Text = OffsetX
                'TextBox29.Text = OffsetY
                Exit Do
            End If
        Loop
        FileClose(100)

        '   put the rest in listbox3

        FileOpen(100, InputPath & InputFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Do Until EOF(100)
            InputString = LineInput(100)
            If InStr(InputString, "<footprint") > 0 Or _
                InStr(InputString, "<trace") > 0 Or _
                InStr(InputString, "<pad") > 0 Then
                InputString = Mid(InputString, 2, InputString.Length)
                ListBox3.Items.Add(InputString)
                sp = Split(InputString, Chr(9))
                'InputString = sp(1)
                Do Until EOF(100)
                    InputString = LineInput(100)
                    InputString = Mid(InputString, 2, InputString.Length)
                    Try
                        'InputString = sp(1) & sp(2)
                        ListBox3.Items.Add(InputString)
                    Catch ex As Exception
                    End Try
                Loop

            End If
        Loop
        FileClose(100)

        ''ListBox3.Items.RemoveAt(ListBox3.Items.Count - 1)
        '********************************************************************
        '********************************************************************


        'Exit Sub
        '********************************************************************
        '********************************************************************
        ''InputPath = Mid(CurDir, 1, Len(CurDir) - 8) & "Stripboards\Footprints\"
        InputPath = BasePath & "\Footprints\"



        '   load in the footprints

        FileOpen(100, InputPath & "Footprints 01.txt", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Do Until EOF(100)
            InputString = LineInput(100)
            'ListBox19.Items.Add(InputString)
            If InStr(InputString, "<footprint = ") > 0 Then
                sp = Split(InputString, Chr(34))
                ComboBox7.Items.Add(sp(1))
            End If
        Loop
        ComboBox7.Text = "component footprints"
        FileClose(100)
        '********************************************************************
        '********************************************************************
        FileClose()

        ListBox20.Items.Clear()

        XOrigin = OffsetX
        YOrigin = OffsetY

        x1 = OffsetX
        y1 = 1500 - OffsetY

        XDisplay1 = OffsetX / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY) / DisplayScale
        XDisplay2 = OffsetX / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        XDisplay1 = OffsetX / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale
        XDisplay2 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        XDisplay1 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale
        XDisplay2 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        XDisplay1 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY) / DisplayScale
        XDisplay2 = OffsetX / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        '********************************************************************
        '********************************************************************
        Dim x2, y2 As Single
        Dim x20, y20 As Single
        Dim x20first, y20first As Single
        Dim x20last, y20last As Single
        '********************************************************************
        '********************************************************************
        'Dim TrimmedString As String

        For ii = 0 To ListBox3.Items.Count - 1
            If InStr(ListBox3.Items.Item(ii), "<footprint = ") > 0 Then

                '   get the location

                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))

                x2 = CSng(sp(1))
                y2 = CSng(sp(3))

                '   find the footprint

                FileOpen(100, InputPath & "Footprints 01.txt", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
                Do Until EOF(100)
                    InputString = LineInput(100)
                    If InStr(InputString, ListBox3.Items.Item(ii)) > 0 Then
                        Do Until InStr(InputString, "</strip>") > 0
                            InputString = LineInput(100)
                            If InStr(InputString, "<strip>") > 0 Then

                                sp = Split(InputString, Chr(34))
                                x20first = x2 + CSng(sp(1)) * PinSpacing ' / DisplayScale
                                y20first = y2 + CSng(sp(3)) * PinSpacing ' / DisplayScale
                                x20last = x2 + CSng(sp(5)) * PinSpacing ' / DisplayScale
                                y20last = y2 + CSng(sp(7)) * PinSpacing ' / DisplayScale

                                XDisplay1 = (OffsetX + (x20first - 5)) / DisplayScale
                                YDisplay1 = (y20first + OffsetY - (iii)) / DisplayScale
                                XDisplay2 = (OffsetX + (x20last + 5)) / DisplayScale
                                YDisplay2 = (y20last + OffsetY - (iii)) / DisplayScale

                                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "DarkGold", bmp)
                            End If

                        Loop
                    End If
                Loop
                FileClose(100)
            End If
            sp = Split(ListBox3.Items.Item(ii), Chr(34))
            PenWidth = lastPenWidth
        Next

        '********************************************************************
        '********************************************************************
        '********************************************************************
        '********************************************************************


        For ii = PinSpacing To XBoardLimit - PinSpacing Step PinSpacing
            For iii = PinSpacing To YBoardLimit - PinSpacing Step PinSpacing
                XDisplay1 = (OffsetX + (ii - 5)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - (iii)) / DisplayScale
                XDisplay2 = (OffsetX + (ii + 5)) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - (iii)) / DisplayScale

                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Green", bmp)

                XDisplay1 = (OffsetX + (ii)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - (iii - 5)) / DisplayScale
                XDisplay2 = (OffsetX + (ii)) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - (iii + 5)) / DisplayScale

                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Green", bmp)

                XDisplay1 = (OffsetX + (ii)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - (iii)) / DisplayScale
                ListBox20.Items.Add(Str(ii) & Chr(9) & Str(iii) & Chr(9) & Str(XDisplay1) & Chr(9) & Str(YDisplay1))

            Next
        Next

        PictureBox1.Image = bmp
        graphics.Dispose()
        PictureBox1.Refresh()
        'Exit Sub
        'GoTo skip01
        '********************************************************************
        '********************************************************************
        '********************************************************************
        '********************************************************************

        'REDRAW THE DESIGN

        For ii = 0 To ListBox3.Items.Count - 1
            If InStr(ListBox3.Items.Item(ii), "<footprint = ") > 0 Then

                '   get the location

                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))

                x2 = CSng(sp(1))
                y2 = CSng(sp(3))

                '   find the footprint

                FileOpen(100, InputPath & "Footprints 01.txt", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
                Do Until EOF(100)
                    InputString = LineInput(100)
                    If InStr(InputString, ListBox3.Items.Item(ii)) > 0 Then
                        Do Until InStr(InputString, "</footprint>") > 0
                            InputString = LineInput(100)
                            If InStr(InputString, "<perimeter>") > 0 Then
                                'Dim PerimeterBegin As Boolean = False
                                Do Until InStr(InputString, "</perimeter>") > 0
                                    InputString = LineInput(100)
                                    If InStr(InputString, "</perimeter>") = 0 Then
                                        sp = Split(InputString, Chr(34))
                                        x20first = x2 + CSng(sp(1)) * PinSpacing ' / DisplayScale
                                        y20first = y2 + CSng(sp(3)) * PinSpacing ' / DisplayScale
                                        x20last = x2 + CSng(sp(5)) * PinSpacing ' / DisplayScale
                                        y20last = y2 + CSng(sp(7)) * PinSpacing ' / DisplayScale

                                        XDisplay1 = (OffsetX + x20first) / DisplayScale
                                        YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                                        XDisplay2 = (OffsetX + x20last) / DisplayScale
                                        YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale
                                        If ShowIndex < 2 Then
                                            'Call NewDrawBlueLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
                                            Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Blue", bmp)
                                        End If
                                    End If
                                Loop
                            ElseIf InStr(InputString, "<pins>") > 0 Then
                                Do Until InStr(InputString, "</pins>") > 0
                                    InputString = LineInput(100)
                                    If InStr(InputString, "</pins>") = 0 Then
                                        sp = Split(InputString, Chr(34))
                                        x20 = x2 + CSng(sp(3)) * PinSpacing '/ DisplayScale
                                        y20 = y2 + CSng(sp(5)) * PinSpacing '/ DisplayScale

                                        x20first = x20 - 10 '/ 2
                                        y20first = y20
                                        'sp = Split(ListBox21.Items.Item(iii), Chr(9))
                                        x20last = x20 + 10 '/ 2
                                        y20last = y20
                                        lastPenWidth = PenWidth
                                        PenWidth = 10 'CSng(TextBox24.Text)
                                        XDisplay1 = (OffsetX + x20first) / DisplayScale
                                        YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                                        XDisplay2 = (OffsetX + x20last) / DisplayScale
                                        YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale
                                        'Call NewDrawGoldLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
                                        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Gold", bmp)
                                        PenWidth = lastPenWidth

                                        XDisplay1 = (OffsetX + x20) / DisplayScale
                                        YDisplay1 = (YBoardLimit + OffsetY - y20) / DisplayScale
                                        ''Call NewDrawRedArcLineSegment(XDisplay1, YDisplay1, bmp)

                                        Call DrawCircle(XDisplay2, YDisplay2, "Red", bmp)
                                    End If
                                Loop
                            End If

                        Loop
                    End If
                Loop
                FileClose(100)
            ElseIf InStr(ListBox3.Items.Item(ii), "<trace>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                XDisplay1 = (OffsetX + x20first) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (OffsetX + x20last) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                'Call NewDrawGoldLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)

                'Call NewDrawGoldArcLineSegment(XDisplay1, YDisplay1, bmp)
                'Call NewDrawGoldArcLineSegment(XDisplay2, YDisplay2, bmp)

                Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "Gold", bmp)

                Call DrawCircle(XDisplay1, YDisplay1, "Gold", bmp)
                Call DrawCircle(XDisplay2, YDisplay2, "Gold", bmp)


            ElseIf InStr(ListBox3.Items.Item(ii), "<pad>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                XDisplay1 = (OffsetX + x20first) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (OffsetX + x20last) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                'Call NewDrawGoldLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Gold", bmp)
            End If
            sp = Split(ListBox3.Items.Item(ii), Chr(34))
            PenWidth = lastPenWidth
        Next
skip01:
        '********************************************************************
        '********************************************************************
        '********************************************************************
        '********************************************************************
        PictureBox1.Image = bmp
        graphics.Dispose()
        PictureBox1.Refresh()
    End Sub
    Private Function LookBeyondTheKnot(ByVal X As Integer, ByVal Y As Integer, ByVal width As Integer, ByRef XFound As Integer, ByRef YFound As Integer, ByRef Listbox23Index As Integer) As Boolean

        LookBeyondTheKnot = False

        Dim i, ii, j, jj As Integer
        Dim LocationX, LocationY As Integer
        Dim iLoops(7), iConnections(7) As Integer
        Dim PathsCounter As Integer

        Dim spi0, spi1 As Object

        Dim HitFlag As Boolean = False

        If X = 132 And Y = 85 Then
            'Stop
        End If

        '   do the top

        For i = X - width To X + width

            For j = 0 To ListBox23.Items.Count - 2 Step 2

                spi0 = Split(ListBox23.Items.Item(j), "|")

                LocationX = CInt(spi0(0))
                LocationY = CInt(spi0(1))

                If LocationX = 85 And LocationY = 144 Then
                    'Stop
                End If

                If LocationX = i And LocationY = Y + width Then
                    spi0 = Split(ListBox23.Items.Item(j), "|")
                    spi1 = Split(ListBox23.Items.Item(j + 1), "|")
                    For ii = 0 To 7 Step 1
                        iConnections(ii) = CInt(spi0(ii + 3))
                        iLoops(ii) = CInt(spi1(ii + 1))
                    Next ii

                    PathsCounter = 0

                    For jj = 0 To 7
                        If iConnections(jj) = 1 And iLoops(jj) = 0 And DevelopSliceActiveFlag(j) = True Then
                            PathsCounter = PathsCounter + 1

                            If HitFlag = False Then
                                XFound = LocationX
                                YFound = LocationY
                                LookBeyondTheKnot = True
                                Listbox23Index = j
                                'ListBox19.Items.Add(Chr(9) & Chr(9) & "TOP    " & Str(LocationX) & Chr(9) & Str(LocationY))
                                HitFlag = True
                                Grid(LocationX, LocationY) = 10
                            Else
                                If Grid(LocationX, LocationY) <> 10 Then
                                    DevelopSliceActiveFlag(j) = False
                                    DevelopSliceActiveFlag(j + 1) = False
                                    Grid(LocationX, LocationY) = 0
                                End If
                            End If

                            TextBox17.Text = Str(LocationX)
                            TextBox17.Refresh()
                            TextBox18.Text = Str(LocationY)
                            TextBox18.Refresh()

                        End If
                    Next jj
                End If
            Next
        Next

        '   do the bottom

        For i = X + width To X - width Step -1

            For j = 0 To ListBox23.Items.Count - 2 Step 2

                spi0 = Split(ListBox23.Items.Item(j), "|")

                LocationX = CInt(spi0(0))
                LocationY = CInt(spi0(1))

                If LocationX = 85 And LocationY = 144 Then
                    'Stop
                End If

                If LocationX = i And LocationY = Y - width Then
                    spi0 = Split(ListBox23.Items.Item(j), "|")
                    spi1 = Split(ListBox23.Items.Item(j + 1), "|")
                    For ii = 0 To 7 Step 1
                        iConnections(ii) = CInt(spi0(ii + 3))
                        iLoops(ii) = CInt(spi1(ii + 1))
                    Next ii

                    PathsCounter = 0

                    For jj = 0 To 7
                        If iConnections(jj) = 1 And iLoops(jj) = 0 And DevelopSliceActiveFlag(j) = True Then
                            PathsCounter = PathsCounter + 1
                            If HitFlag = False Then
                                XFound = LocationX
                                YFound = LocationY
                                LookBeyondTheKnot = True
                                Listbox23Index = j
                                'ListBox19.Items.Add(Chr(9) & Chr(9) & "BOTTOM " & Str(LocationX) & Chr(9) & Str(LocationY))
                                HitFlag = True
                                Grid(LocationX, LocationY) = 10
                            Else
                                If Grid(LocationX, LocationY) <> 10 Then
                                    DevelopSliceActiveFlag(j) = False
                                    DevelopSliceActiveFlag(j + 1) = False
                                    Grid(LocationX, LocationY) = 0
                                End If
                            End If
                            TextBox17.Text = Str(LocationX)
                            TextBox17.Refresh()
                            TextBox18.Text = Str(LocationY)
                            TextBox18.Refresh()
                        End If
                    Next jj
                End If
            Next
        Next


        '   do the right side

        For i = Y + width - 1 To Y - width + 1 Step -1

            For j = 0 To ListBox23.Items.Count - 2 Step 2

                spi0 = Split(ListBox23.Items.Item(j), "|")

                LocationX = CInt(spi0(0))
                LocationY = CInt(spi0(1))

                TextBox17.Text = Str(LocationX)
                TextBox18.Text = Str(LocationX)

                If LocationX = 85 And LocationY = 144 Then
                    'Stop
                End If

                If LocationX = X + width And LocationY = i Then
                    spi0 = Split(ListBox23.Items.Item(j), "|")
                    spi1 = Split(ListBox23.Items.Item(j + 1), "|")
                    For ii = 0 To 7 Step 1
                        iConnections(ii) = CInt(spi0(ii + 3))
                        iLoops(ii) = CInt(spi1(ii + 1))
                    Next ii

                    PathsCounter = 0

                    For jj = 0 To 7
                        If iConnections(jj) = 1 And iLoops(jj) = 0 And DevelopSliceActiveFlag(j) = True Then
                            PathsCounter = PathsCounter + 1
                            If HitFlag = False Then
                                XFound = LocationX
                                YFound = LocationY
                                LookBeyondTheKnot = True
                                Listbox23Index = j
                                'ListBox19.Items.Add(Chr(9) & Chr(9) & "RIGHT  " & Str(LocationX) & Chr(9) & Str(LocationY))
                                HitFlag = True
                                Grid(LocationX, LocationY) = 10
                            Else
                                If Grid(LocationX, LocationY) <> 10 Then
                                    DevelopSliceActiveFlag(j) = False
                                    DevelopSliceActiveFlag(j + 1) = False
                                    Grid(LocationX, LocationY) = 0
                                End If
                            End If
                            TextBox17.Text = Str(LocationX)
                            TextBox17.Refresh()
                            TextBox18.Text = Str(LocationY)
                            TextBox18.Refresh()
                        End If
                    Next jj
                End If
            Next
        Next


        '   do the left side

        For i = Y - width + 1 To Y + width - 1

            For j = 0 To ListBox23.Items.Count - 2 Step 2

                spi0 = Split(ListBox23.Items.Item(j), "|")

                LocationX = CInt(spi0(0))
                LocationY = CInt(spi0(1))

                If LocationX = 85 And LocationY = 144 Then
                    'Stop
                End If

                If LocationX = X - width And LocationY = i Then
                    spi0 = Split(ListBox23.Items.Item(j), "|")
                    spi1 = Split(ListBox23.Items.Item(j + 1), "|")
                    For ii = 0 To 7 Step 1
                        iConnections(ii) = CInt(spi0(ii + 3))
                        iLoops(ii) = CInt(spi1(ii + 1))
                    Next ii

                    PathsCounter = 0

                    For jj = 0 To 7
                        If iConnections(jj) = 1 And iLoops(jj) = 0 And DevelopSliceActiveFlag(j) = True Then
                            PathsCounter = PathsCounter + 1
                            If HitFlag = False Then
                                XFound = LocationX
                                YFound = LocationY
                                LookBeyondTheKnot = True
                                Listbox23Index = j
                                'ListBox19.Items.Add(Chr(9) & Chr(9) & "LEFT   " & Str(LocationX) & Chr(9) & Str(LocationY))
                                HitFlag = True
                                Grid(LocationX, LocationY) = 10
                            Else
                                If Grid(LocationX, LocationY) <> 10 Then
                                    DevelopSliceActiveFlag(j) = False
                                    DevelopSliceActiveFlag(j + 1) = False
                                    Grid(LocationX, LocationY) = 0
                                End If
                            End If
                            'TextBox17.Text = Str(LocationX)
                            'TextBox17.Refresh()
                            'TextBox18.Text = Str(LocationY)
                            'TextBox18.Refresh()
                            'Call PaintGrid()
                            Exit Function
                        End If
                    Next jj
                End If
            Next
        Next

        'Call PaintGrid()

    End Function


    Private Sub Button43_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button43.Click
        Dim bmp As New Bitmap(PictureBox1.Image)
        Dim graphics As Graphics = graphics.FromImage(bmp)

        Button4.Text = "Front"

        graphics.Clear(Color.Ivory)



        PictureBox1.Image = bmp
        graphics.Dispose()
        PictureBox1.Refresh()

        '********************************************************************

        OpenFileDialog1.Filter = "Board Files (*.brd)|*.brd"
        Try
            ''MkDir("C:\Program Files\Clanking Replicators\Stripboard Designer\Boards")
        Catch ex As Exception

        End Try
        OpenFileDialog1.InitialDirectory = BasePath & "\Boards"
        OpenFileDialog1.ShowDialog()

        GroupBox4.Visible = True
        GroupBox5.Visible = True
        GroupBox3.Visible = True
        GroupBox7.Visible = True

        Button8.Visible = True
        Button2.Visible = True

        '********************************************************************
        Dim InputPath As String = BasePath & "\Boards\"
        Dim InputFileName As String = OpenFileDialog1.FileName
        Dim InputString As String


        PenWidth = CSng(TextBox9.Text)


        Dim XDisplay1 As Single
        Dim YDisplay1 As Single

        Dim XDisplay2 As Single
        Dim YDisplay2 As Single
        Dim sp As Object

        Dim BoundariesFlag As Boolean = False

        Dim ii As Single
        Dim iii As Single
        'Dim j As Integer

        Dim x1, y1 As Integer

        Dim LoopCount As Integer = 0
        Dim TraceCount1 As Integer = 0
        Dim TraceCount2 As Integer = 0

        Dim EndsRho(3) As Double
        Dim lastPenWidth As Object

        Dim RhoLimit As Double = 1 * 2 ^ 0.5

        '********************************************************************
        '********************************************************************
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Clear()
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox19.Items.Clear()
        ListBox21.Items.Clear()

        Dim XOrigin, YOrigin As Double

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ComboBox7.Items.Clear()
        '********************************************************************
        '********************************************************************

        '   load in the footprints

        '   get the display scale

        FileOpen(100, InputFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Do Until EOF(100)
            InputString = LineInput(100)
            If InStr(InputString, "<DisplayScale = ") > 0 Then
                sp = Split(InputString, Chr(34))
                DisplayScale = CSng(TextBox35.Text) 'CSng(sp(1))
                'TextBox30.Text = 1 / DisplayScale
                Exit Do
            End If
        Loop
        FileClose(100)

        '   get the pin spacing

        FileOpen(100, InputFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Do Until EOF(100)
            InputString = LineInput(100)
            If InStr(InputString, "<PinSpacing") > 0 Then
                sp = Split(InputString, Chr(34))
                PinSpacing = CSng(sp(1))
                Exit Do
            End If
        Loop
        FileClose(100)

        '   get the board dimensions

        FileOpen(100, InputFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Do Until EOF(100)
            InputString = LineInput(100)
            If InStr(InputString, "<Board dimensions") > 0 Then
                sp = Split(InputString, Chr(34))
                XBoardLimit = CSng(sp(1))
                YBoardLimit = CSng(sp(3))
                TextBox7.Text = CInt(XBoardLimit / 25.4 - 1)
                TextBox8.Text = CInt(YBoardLimit / 25.4)
                Exit Do
            End If
        Loop
        FileClose(100)

        '   get the offsets

        FileOpen(100, InputFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Do Until EOF(100)
            InputString = LineInput(100)
            If InStr(InputString, "<Display offsets") > 0 Then
                sp = Split(InputString, Chr(34))
                OffsetX = 50 'CSng(sp(1))
                OffsetY = 50 'CSng(sp(3))
                'TextBox28.Text = OffsetX
                'TextBox29.Text = OffsetY
                Exit Do
            End If
        Loop
        FileClose(100)

        '   put the rest in listbox3

        FileOpen(100, InputFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Do Until EOF(100)
            InputString = LineInput(100)
            If InStr(InputString, "<footprint") > 0 Or _
                InStr(InputString, "<trace") > 0 Or _
                InStr(InputString, "<break") > 0 Or _
                InStr(InputString, "<stripboard jumper") > 0 Or _
                InStr(InputString, "<pad") > 0 Then
                InputString = Mid(InputString, 2, InputString.Length)
                ListBox3.Items.Add(InputString)
                sp = Split(InputString, Chr(9))
                'InputString = sp(1)
                Do Until EOF(100)
                    InputString = LineInput(100)
                    InputString = Mid(InputString, 2, InputString.Length)
                    Try
                        'InputString = sp(1) & sp(2)
                        ListBox3.Items.Add(InputString)
                    Catch ex As Exception
                    End Try
                Loop

            End If
        Loop
        FileClose(100)

        'Exit Sub
        Try
            ListBox3.Items.RemoveAt(ListBox3.Items.Count - 1)
        Catch ex As Exception

        End Try

        '********************************************************************
        '********************************************************************


        'Exit Sub
        '********************************************************************
        '********************************************************************
        Dim FootprintPath As String = BasePath & "\Footprints\"


        '   load in the footprints

        FileOpen(100, FootprintPath & "Footprints 01.txt", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Do Until EOF(100)
            InputString = LineInput(100)
            'ListBox19.Items.Add(InputString)
            If InStr(InputString, "<footprint = ") > 0 Then
                sp = Split(InputString, Chr(34))
                ComboBox7.Items.Add(sp(1))
            End If
        Loop
        ComboBox7.Text = "component footprints"
        FileClose(100)
        '********************************************************************
        '********************************************************************
        FileClose()

        ListBox20.Items.Clear()

        XOrigin = OffsetX
        YOrigin = OffsetY

        x1 = OffsetX
        y1 = 1500 - OffsetY

        XDisplay1 = OffsetX / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY + PinSpacing) / DisplayScale
        XDisplay2 = OffsetX / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        XDisplay1 = OffsetX / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale
        XDisplay2 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        XDisplay1 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale
        XDisplay2 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY + PinSpacing) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        XDisplay1 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY + PinSpacing) / DisplayScale
        XDisplay2 = OffsetX / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY + PinSpacing) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        'PictureBox1.Image = bmp
        'graphics.Dispose()
        'PictureBox1.Refresh()
        'Exit Sub

        PenWidth = 9

        Dim ii1 As Integer = 0.75 * PinSpacing
        Dim ii2 As Integer = XBoardLimit - 0.75 * PinSpacing

        ''For ii = PinSpacing To XBoardLimit - PinSpacing Step XBoardLimit - 2 * PinSpacing
        'For iii = PinSpacing To YBoardLimit - PinSpacing Step PinSpacing
        For iii = 1 To YBoardLimit Step PinSpacing
            XDisplay1 = (OffsetX + (ii1 - 5)) / DisplayScale
            YDisplay1 = (YBoardLimit + OffsetY - (iii)) / DisplayScale
            XDisplay2 = (OffsetX + (ii2 + 5)) / DisplayScale
            YDisplay2 = (YBoardLimit + OffsetY - (iii)) / DisplayScale

            Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "DarkGold", bmp)

            ListBox3.Items.Add("<strip>")
            ListBox3.Items.Add(Chr(9) & "<penwidth =" & Chr(34) & Str(PenWidth) & Chr(34) & "/>")
            ListBox3.Items.Add(Chr(9) & "<begin x=" & Chr(34) & Str(ii1) & Chr(34) & " y=" & Chr(34) & Str(iii) & Chr(34) & "/>")
            ListBox3.Items.Add(Chr(9) & "<end   x=" & Chr(34) & Str(ii2) & Chr(34) & " y=" & Chr(34) & Str(iii) & Chr(34) & "/>")
            ListBox3.Items.Add("</strip>")


            XDisplay1 = (OffsetX + (ii)) / DisplayScale
            YDisplay1 = (YBoardLimit + OffsetY - (iii - 5)) / DisplayScale
            XDisplay2 = (OffsetX + (ii)) / DisplayScale
            YDisplay2 = (YBoardLimit + OffsetY - (iii + 5)) / DisplayScale

            'Call NewDrawGoldLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)


            XDisplay1 = (ii) / DisplayScale + OffsetX
            YDisplay1 = YBoardLimit + OffsetY - (iii) / DisplayScale
            ListBox20.Items.Add(Str(ii) & Chr(9) & Str(iii) & Chr(9) & Str(XDisplay1) & Chr(9) & Str(YDisplay1))

        Next
        ''Next

        PenWidth = 5

        'For ii = PinSpacing To XBoardLimit - PinSpacing Step PinSpacing
        For ii = PinSpacing To XBoardLimit Step PinSpacing
            'For iii = PinSpacing To YBoardLimit - PinSpacing Step PinSpacing
            For iii = 1 To YBoardLimit Step PinSpacing

                XDisplay1 = (OffsetX + (ii - 5)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - (iii)) / DisplayScale
                XDisplay2 = (OffsetX + (ii + 5)) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - (iii)) / DisplayScale

                'Call NewDrawGreenLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)

                XDisplay1 = (OffsetX + (ii)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - (iii)) / DisplayScale

                ''Call NewDrawBlackArcLineSegment(XDisplay1, YDisplay1, bmp)

                Call DrawCircle(XDisplay1, YDisplay1, "Black", bmp)

                XDisplay1 = (OffsetX + (ii)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - (iii - 5)) / DisplayScale
                XDisplay2 = (OffsetX + (ii)) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - (iii + 5)) / DisplayScale

                'Call NewDrawGreenLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)


                XDisplay1 = (ii) / DisplayScale + OffsetX
                YDisplay1 = YBoardLimit + OffsetY - (iii) / DisplayScale
                ListBox20.Items.Add(Str(ii) & Chr(9) & Str(iii) & Chr(9) & Str(XDisplay1) & Chr(9) & Str(YDisplay1))

            Next
        Next

        PenWidth = CSng(TextBox9.Text)

        PictureBox1.Image = bmp
        graphics.Dispose()
        PictureBox1.Refresh()

        PenWidth = 3
        'Exit Sub
        'GoTo skip01
        '********************************************************************
        '********************************************************************
        Dim x2, y2 As Single
        Dim x20, y20 As Single
        Dim x20first, y20first As Single
        Dim x20last, y20last As Single
        '********************************************************************
        '********************************************************************
        'Dim TrimmedString As String

        'REDRAW THE DESIGN

        For ii = 0 To ListBox3.Items.Count - 1
            If InStr(ListBox3.Items.Item(ii), "<footprint = ") > 0 Then

                '   get the location

                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))

                x2 = CSng(sp(1))
                y2 = CSng(sp(3))

                '   find the footprint
                InputPath = BasePath & "\Footprints\"

                FileOpen(100, InputPath & "Footprints 01.txt", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
                Do Until EOF(100)
                    InputString = LineInput(100)
                    If InStr(InputString, ListBox3.Items.Item(ii)) > 0 Then
                        Do Until InStr(InputString, "</footprint>") > 0
                            InputString = LineInput(100)
                            If InStr(InputString, "<perimeter>") > 0 Then
                                'Dim PerimeterBegin As Boolean = False
                                Do Until InStr(InputString, "</perimeter>") > 0
                                    InputString = LineInput(100)
                                    If InStr(InputString, "</perimeter>") = 0 Then
                                        sp = Split(InputString, Chr(34))
                                        x20first = x2 + CSng(sp(1)) * PinSpacing ' / DisplayScale
                                        y20first = y2 + CSng(sp(3)) * PinSpacing ' / DisplayScale
                                        x20last = x2 + CSng(sp(5)) * PinSpacing ' / DisplayScale
                                        y20last = y2 + CSng(sp(7)) * PinSpacing ' / DisplayScale

                                        XDisplay1 = (OffsetX + x20first) / DisplayScale
                                        YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                                        XDisplay2 = (OffsetX + x20last) / DisplayScale
                                        YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale
                                        If ShowIndex < 2 Then
                                            'Call NewDrawBlueLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
                                            Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Blue", bmp)
                                        End If
                                    End If
                                Loop
                            ElseIf InStr(InputString, "<pins>") > 0 Then
                                Do Until InStr(InputString, "</pins>") > 0
                                    InputString = LineInput(100)
                                    If InStr(InputString, "</pins>") = 0 Then
                                        sp = Split(InputString, Chr(34))
                                        x20 = x2 + CSng(sp(3)) * PinSpacing '/ DisplayScale
                                        y20 = y2 + CSng(sp(5)) * PinSpacing '/ DisplayScale

                                        x20first = x20 - 10 '/ 2
                                        y20first = y20
                                        'sp = Split(ListBox21.Items.Item(iii), Chr(9))
                                        x20last = x20 + 10 '/ 2
                                        y20last = y20
                                        lastPenWidth = PenWidth
                                        PenWidth = 10 'CSng(TextBox24.Text)
                                        XDisplay1 = (OffsetX + x20first) / DisplayScale
                                        YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                                        XDisplay2 = (OffsetX + x20last) / DisplayScale
                                        YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale
                                        'Call NewDrawGoldLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
                                        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Gold", bmp)
                                        PenWidth = lastPenWidth

                                        XDisplay1 = (OffsetX + x20) / DisplayScale
                                        YDisplay1 = (YBoardLimit + OffsetY - y20) / DisplayScale
                                        ''Call NewDrawRedArcLineSegment(XDisplay1, YDisplay1, bmp)

                                        Call DrawCircle(XDisplay1, YDisplay1, "Red", bmp)
                                    End If
                                Loop
                            End If

                        Loop
                    End If
                Loop
                FileClose(100)
            ElseIf InStr(ListBox3.Items.Item(ii), "<break>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                XDisplay1 = (OffsetX + x20first) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (OffsetX + x20last) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                'Call NewDrawIvoryLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Ivory", bmp)
            ElseIf InStr(ListBox3.Items.Item(ii), "<twixt break>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                XDisplay1 = (OffsetX + x20first) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (OffsetX + x20last) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                'Call NewDrawIvoryLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Ivory", bmp)
            ElseIf InStr(ListBox3.Items.Item(ii), "<stripboard jumper>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                XDisplay1 = (OffsetX + x20first) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (OffsetX + x20last) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                'Call NewDrawOrangeLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "HotPink", bmp)

                Call DrawCircle(XDisplay1, YDisplay1, "HotPink", bmp)
                Call DrawCircle(XDisplay2, YDisplay2, "HotPink", bmp)
            ElseIf InStr(ListBox3.Items.Item(ii), "<stripboard jumper - gnd>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                XDisplay1 = (OffsetX + x20first) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (OffsetX + x20last) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                ''Call NewDrawGreenLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)

                ''Call NewDrawGreenArcLineSegment(XDisplay1, YDisplay1, bmp)
                ''Call NewDrawGreenArcLineSegment(XDisplay2, YDisplay2, bmp)

                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Green", bmp)

                Call DrawCircle(XDisplay1, YDisplay1, "Green", bmp)
                Call DrawCircle(XDisplay2, YDisplay2, "Green", bmp)

            ElseIf InStr(ListBox3.Items.Item(ii), "<stripboard jumper - +5 v>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                XDisplay1 = (OffsetX + x20first) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (OffsetX + x20last) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "DarkRed", bmp)

                Call DrawCircle(XDisplay1, YDisplay1, "DarkRed", bmp)
                Call DrawCircle(XDisplay2, YDisplay2, "DarkRed", bmp)

            ElseIf InStr(ListBox3.Items.Item(ii), "<stripboard jumper - +12 v>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                XDisplay1 = (OffsetX + x20first) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (OffsetX + x20last) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                'Call NewDrawRedLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)

                'Call NewDrawRedArcLineSegment(XDisplay1, YDisplay1, bmp)
                'Call NewDrawRedArcLineSegment(XDisplay2, YDisplay2, bmp)

                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

                Call DrawCircle(XDisplay1, YDisplay1, "Red", bmp)
                Call DrawCircle(XDisplay2, YDisplay2, "Red", bmp)


            ElseIf InStr(ListBox3.Items.Item(ii), "<stripboard jumper - 0 - +5 v>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                XDisplay1 = (OffsetX + x20first) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (OffsetX + x20last) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                'Call NewDrawBlackLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)

                'Call NewDrawBlackArcLineSegment(XDisplay1, YDisplay1, bmp)
                'Call NewDrawBlackArcLineSegment(XDisplay2, YDisplay2, bmp)

                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Black", bmp)

                Call DrawCircle(XDisplay1, YDisplay1, "Black", bmp)
                Call DrawCircle(XDisplay2, YDisplay2, "Black", bmp)

            ElseIf InStr(ListBox3.Items.Item(ii), "<stripboard jumper - 0 - +12 v>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                XDisplay1 = (OffsetX + x20first) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (OffsetX + x20last) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                'Call NewDrawPurpleLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)

                'Call NewDrawPurpleArcLineSegment(XDisplay1, YDisplay1, bmp)
                'Call NewDrawPurpleArcLineSegment(XDisplay2, YDisplay2, bmp)

                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Purple", bmp)

                Call DrawCircle(XDisplay1, YDisplay1, "Purple", bmp)
                Call DrawCircle(XDisplay2, YDisplay2, "Purple", bmp)
            End If
            sp = Split(ListBox3.Items.Item(ii), Chr(34))
            PenWidth = lastPenWidth
        Next
skip01:

        PenWidth = CSng(TextBox9.Text)
        '********************************************************************
        '********************************************************************
        '********************************************************************
        '********************************************************************
        PictureBox1.Image = bmp
        graphics.Dispose()
        PictureBox1.Refresh()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim ii As Integer

        listbox5.Items.Clear()


        listbox5.Items.Add("<Stripboard>")
        listbox5.Items.Add(Chr(9) & "<DisplayScale = " & Chr(34) & Str(DisplayScale) & Chr(34) & "/>")
        listbox5.Items.Add(Chr(9) & "<PinSpacing (mm/10) = " & Chr(34) & Str(PinSpacing) & Chr(34) & "/>")
        listbox5.Items.Add(Chr(9) & "<Board dimensions (mm/10) x = " & Chr(34) & Str(XBoardLimit) & Chr(34) & _
                                    " y = " & Chr(34) & Str(YBoardLimit) & Chr(34) & "/>")
        listbox5.Items.Add(Chr(9) & "<Display offsets x = " & Chr(34) & Str(OffsetX) & Chr(34) & _
                            " y = " & Chr(34) & Str(OffsetY) & Chr(34) & "/>")

        For ii = 0 To ListBox3.Items.Count - 1
            listbox5.Items.Add(Chr(9) & ListBox3.Items.Item(ii))
        Next

        listbox5.Items.Add("</Stripboard>")
        '********************************************************************

        SaveFileDialog1.Filter = "Board Files (*.brd)|*.brd"
        Try
            ''MkDir("C:\Program Files\Clanking Replicators\Stripboard Designer\Boards")
        Catch ex As Exception

        End Try
        SaveFileDialog1.InitialDirectory = BasePath & "\Boards"
        SaveFileDialog1.ShowDialog()

        '********************************************************************
        Dim OutputPath As String = BasePath & "\Boards\"
        Dim OutputFileName As String = SaveFileDialog1.FileName
        If InStr(OutputFileName, ".brd") = 0 Then
            OutputFileName = OutputFileName & ".brd"
        End If

        '   load in the footprints

        FileOpen(100, OutputFileName, OpenMode.Output, OpenAccess.Write, OpenShare.Shared)

        For ii = 0 To listbox5.Items.Count - 1
            PrintLine(100, listbox5.Items.Item(ii))
        Next
        FileClose(100)
        Exit Sub
        '**********************************************************************************
        '**********************************************************************************
        '**********************************************************************************
        '**********************************************************************************

        '**********************************************************************************
        '**********************************************************************************
        '**********************************************************************************
        '**********************************************************************************
    End Sub

    Private Sub Button41_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button41.Click
        ToolFlag = "break"
        PenWidth = CInt(TextBox24.Text)
        ListBox21.Items.Clear()
    End Sub

    Private Sub Button42_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button42.Click
        ToolFlag = "stripboard jumper"
        PenWidth = CInt(TextBox23.Text)
        ListBox21.Items.Clear()
    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        ToolFlag = "select"
        'PenWidth = CInt(TextBox23.Text)
        'ListBox21.Items.Clear()
    End Sub

    Private Sub RefreshStripboardDisplay(ByVal ShowIndex As Integer)
        Dim bmp As New Bitmap(PictureBox1.Image)
        Dim graphics As Graphics = graphics.FromImage(bmp)

        graphics.Clear(Color.Ivory)

        PictureBox1.Image = bmp
        graphics.Dispose()
        PictureBox1.Refresh()

        'Exit Sub

        PenWidth = CSng(TextBox9.Text)


        Dim XDisplay1 As Single
        Dim YDisplay1 As Single

        Dim XDisplay2 As Single
        Dim YDisplay2 As Single
        Dim sp As Object

        Dim BoundariesFlag As Boolean = False

        Dim ii As Single
        Dim iii As Single
        'Dim j As Integer

        Dim x1, y1 As Integer

        Dim LoopCount As Integer = 0
        Dim TraceCount1 As Integer = 0
        Dim TraceCount2 As Integer = 0

        Dim EndsRho(3) As Double

        Dim RhoLimit As Double = 1 * 2 ^ 0.5

        '********************************************************************
        '********************************************************************

        'XBoardLimit = CInt(TextBox26.Text)
        'YBoardLimit = CInt(TextBox27.Text)
        'OffsetX = CInt(TextBox28.Text)
        'OffsetY = CInt(TextBox29.Text)
        DisplayScale = CSng(TextBox35.Text) '1 / CSng(TextBox30.Text)

        PinSpacing = 25.4 '* 10

        PenWidth = CSng(TextBox9.Text)

        '********************************************************************
        '********************************************************************

        Dim lastPenWidth As Single

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        'ListBox3.Items.Clear()
        ListBox4.Items.Clear()
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox19.Items.Clear()
        ListBox21.Items.Clear()

        Dim XOrigin, YOrigin As Double

        DisplayScale = CSng(TextBox35.Text) '1 / CSng(TextBox30.Text)



        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ComboBox7.Items.Clear()

        '********************************************************************
        '********************************************************************
        ''Dim InputPath As String = Mid(CurDir, 1, Len(CurDir) - 8) & "Stripboards\Footprints\"
        Dim InputPath As String = BasePath & "\Footprints\"
        Dim InputString As String

        '   load in the footprints

        FileOpen(100, InputPath & "Footprints 01.txt", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Do Until EOF(100)
            InputString = LineInput(100)
            'ListBox19.Items.Add(InputString)
            If InStr(InputString, "<footprint = ") > 0 Then
                sp = Split(InputString, Chr(34))
                ComboBox7.Items.Add(sp(1))
            End If
        Loop
        ComboBox7.Text = "component footprints"
        FileClose(100)
        '********************************************************************
        '********************************************************************
        FileClose()


        ListBox20.Items.Clear()

        XOrigin = OffsetX
        YOrigin = OffsetY

        x1 = OffsetX
        y1 = YBoardLimit - OffsetY

        XDisplay1 = OffsetX / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY + PinSpacing) / DisplayScale
        XDisplay2 = OffsetX / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        XDisplay1 = OffsetX / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale
        XDisplay2 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        XDisplay1 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale
        XDisplay2 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY + PinSpacing) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        XDisplay1 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY + PinSpacing) / DisplayScale
        XDisplay2 = OffsetX / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY + PinSpacing) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        'PictureBox1.Image = bmp
        'graphics.Dispose()
        'PictureBox1.Refresh()
        'Exit Sub

        PenWidth = 9

        Dim ii1 As Integer = 0.75 * PinSpacing
        Dim ii2 As Integer = XBoardLimit - 0.75 * PinSpacing

        ''For ii = PinSpacing To XBoardLimit - PinSpacing Step XBoardLimit - 2 * PinSpacing
        'For iii = PinSpacing To YBoardLimit - PinSpacing Step PinSpacing
        For iii = 1 To YBoardLimit Step PinSpacing
            XDisplay1 = (OffsetX + (ii1 - 5)) / DisplayScale
            YDisplay1 = (YBoardLimit + OffsetY - (iii)) / DisplayScale
            XDisplay2 = (OffsetX + (ii2 + 5)) / DisplayScale
            YDisplay2 = (YBoardLimit + OffsetY - (iii)) / DisplayScale

            ''Call NewDrawDarkGoldLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
            Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "DarkGold", bmp)

            ListBox3.Items.Add("<strip>")
            ListBox3.Items.Add(Chr(9) & "<penwidth =" & Chr(34) & Str(PenWidth) & Chr(34) & "/>")
            ListBox3.Items.Add(Chr(9) & "<begin x=" & Chr(34) & Str(ii1) & Chr(34) & " y=" & Chr(34) & Str(iii) & Chr(34) & "/>")
            ListBox3.Items.Add(Chr(9) & "<end   x=" & Chr(34) & Str(ii2) & Chr(34) & " y=" & Chr(34) & Str(iii) & Chr(34) & "/>")
            ListBox3.Items.Add("</strip>")


            XDisplay1 = (OffsetX + (ii)) / DisplayScale
            YDisplay1 = (YBoardLimit + OffsetY - (iii - 5)) / DisplayScale
            XDisplay2 = (OffsetX + (ii)) / DisplayScale
            YDisplay2 = (YBoardLimit + OffsetY - (iii + 5)) / DisplayScale

            'Call NewDrawGoldLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)


            XDisplay1 = (ii) / DisplayScale + OffsetX
            YDisplay1 = YBoardLimit + OffsetY - (iii) / DisplayScale
            ListBox20.Items.Add(Str(ii) & Chr(9) & Str(iii) & Chr(9) & Str(XDisplay1) & Chr(9) & Str(YDisplay1))

        Next
        ''Next

        PenWidth = 5

        'For ii = PinSpacing To XBoardLimit - PinSpacing Step PinSpacing
        For ii = PinSpacing To XBoardLimit Step PinSpacing
            'For iii = PinSpacing To YBoardLimit - PinSpacing Step PinSpacing
            For iii = 1 To YBoardLimit Step PinSpacing

                XDisplay1 = (OffsetX + (ii - 5)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - (iii)) / DisplayScale
                XDisplay2 = (OffsetX + (ii + 5)) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - (iii)) / DisplayScale

                'Call NewDrawGreenLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)

                XDisplay1 = (OffsetX + (ii)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - (iii)) / DisplayScale

                ''Call NewDrawBlackArcLineSegment(XDisplay1, YDisplay1, bmp)

                Call DrawCircle(XDisplay1, YDisplay1, "Black", bmp)

                XDisplay1 = (OffsetX + (ii)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - (iii - 5)) / DisplayScale
                XDisplay2 = (OffsetX + (ii)) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - (iii + 5)) / DisplayScale

                'Call NewDrawGreenLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)


                XDisplay1 = (ii) / DisplayScale + OffsetX
                YDisplay1 = YBoardLimit + OffsetY - (iii) / DisplayScale
                ListBox20.Items.Add(Str(ii) & Chr(9) & Str(iii) & Chr(9) & Str(XDisplay1) & Chr(9) & Str(YDisplay1))

            Next
        Next

        PenWidth = CSng(TextBox9.Text)


        'GoTo skip01
        '********************************************************************
        '********************************************************************
        Dim x2, y2 As Single
        Dim x20, y20 As Single
        Dim x20first, y20first As Single
        Dim x20last, y20last As Single
        '********************************************************************
        '********************************************************************

        'REDRAW THE DESIGN

        For ii = 0 To ListBox3.Items.Count - 1
            If InStr(ListBox3.Items.Item(ii), "<footprint = ") > 0 Then

                '   get the location

                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))

                x2 = CSng(sp(1))
                y2 = CSng(sp(3))

                '   find the footprint

                InputPath = BasePath & "\Footprints\"

                FileOpen(100, InputPath & "Footprints 01.txt", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
                Do Until EOF(100)
                    InputString = LineInput(100)
                    If InStr(InputString, ListBox3.Items.Item(ii)) > 0 Then
                        Do Until InStr(InputString, "</footprint>") > 0
                            InputString = LineInput(100)
                            If InStr(InputString, "<perimeter>") > 0 Then
                                'Dim PerimeterBegin As Boolean = False
                                Do Until InStr(InputString, "</perimeter>") > 0
                                    InputString = LineInput(100)
                                    If InStr(InputString, "</perimeter>") = 0 Then
                                        sp = Split(InputString, Chr(34))
                                        x20first = x2 + CSng(sp(1)) * PinSpacing ' / DisplayScale
                                        y20first = y2 + CSng(sp(3)) * PinSpacing ' / DisplayScale
                                        x20last = x2 + CSng(sp(5)) * PinSpacing ' / DisplayScale
                                        y20last = y2 + CSng(sp(7)) * PinSpacing ' / DisplayScale

                                        XDisplay1 = (OffsetX + x20first) / DisplayScale
                                        YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                                        XDisplay2 = (OffsetX + x20last) / DisplayScale
                                        YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                                        If ShowIndex < 2 Then
                                            PenWidth = 3
                                            Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Blue", bmp)
                                            PenWidth = lastPenWidth
                                        End If
                                    End If
                                Loop
                            ElseIf InStr(InputString, "<pins>") > 0 Then
                                Do Until InStr(InputString, "</pins>") > 0
                                    InputString = LineInput(100)
                                    If InStr(InputString, "</pins>") = 0 Then
                                        sp = Split(InputString, Chr(34))
                                        x20 = x2 + CSng(sp(3)) * PinSpacing '/ DisplayScale
                                        y20 = y2 + CSng(sp(5)) * PinSpacing '/ DisplayScale

                                        x20first = x20 - 10 '/ 2
                                        y20first = y20
                                        'sp = Split(ListBox21.Items.Item(iii), Chr(9))
                                        x20last = x20 + 10 '/ 2
                                        y20last = y20
                                        lastPenWidth = PenWidth
                                        PenWidth = 10 'CSng(TextBox24.Text)
                                        XDisplay1 = (OffsetX + x20first) / DisplayScale
                                        YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                                        XDisplay2 = (OffsetX + x20last) / DisplayScale
                                        YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale
                                        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Gold", bmp)
                                        PenWidth = lastPenWidth

                                        XDisplay1 = (OffsetX + x20) / DisplayScale
                                        YDisplay1 = (YBoardLimit + OffsetY - y20) / DisplayScale
                                        If ShowIndex < 2 Then
                                            ''Call NewDrawRedArcLineSegment(XDisplay1, YDisplay1, bmp)

                                            Call DrawCircle(XDisplay1, YDisplay1, "Red", bmp)
                                        End If
                                    End If
                                Loop
                            End If

                        Loop
                    End If
                Loop
                FileClose(100)
            ElseIf InStr(ListBox3.Items.Item(ii), "<stripboard jumper>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                XDisplay1 = (OffsetX + x20first) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (OffsetX + x20last) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                'Call NewDrawOrangeLineSegment(XDisplay1, YDisplay1, XDisplay1, YDisplay2, bmp)

                'Call NewDrawOrangeArcLineSegment(XDisplay1, YDisplay1, bmp)
                'Call NewDrawOrangeArcLineSegment(XDisplay1, YDisplay2, bmp)

                Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "HotPink", bmp)

                Call DrawCircle(XDisplay1, YDisplay1, "HotPink", bmp)
                Call DrawCircle(XDisplay1, YDisplay2, "HotPink", bmp)

            ElseIf InStr(ListBox3.Items.Item(ii), "<stripboard jumper - gnd>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                XDisplay1 = (OffsetX + x20first) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (OffsetX + x20last) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "DarkGreen", bmp)

                Call DrawCircle(XDisplay1, YDisplay1, "DarkGreen", bmp)
                Call DrawCircle(XDisplay1, YDisplay2, "DarkGreen", bmp)

            ElseIf InStr(ListBox3.Items.Item(ii), "<stripboard jumper - +5 v>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                XDisplay1 = (OffsetX + x20first) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (OffsetX + x20last) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                'Call NewDrawDarkRedLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)

                'Call NewDrawGoldArcLineSegment(XDisplay1, YDisplay1, bmp)
                'Call NewDrawGoldArcLineSegment(XDisplay2, YDisplay2, bmp)

                Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "DarkRed", bmp)

                Call DrawCircle(XDisplay1, YDisplay1, "DarkRed", bmp)
                Call DrawCircle(XDisplay1, YDisplay2, "DarkRed", bmp)


            ElseIf InStr(ListBox3.Items.Item(ii), "<stripboard jumper - +12 v>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                XDisplay1 = (OffsetX + x20first) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (OffsetX + x20last) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                'Call NewDrawRedLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)

                'Call NewDrawRedArcLineSegment(XDisplay1, YDisplay1, bmp)
                'Call NewDrawRedArcLineSegment(XDisplay2, YDisplay2, bmp)

                Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "Red", bmp)

                Call DrawCircle(XDisplay1, YDisplay1, "Red", bmp)
                Call DrawCircle(XDisplay1, YDisplay2, "Red", bmp)

            ElseIf InStr(ListBox3.Items.Item(ii), "<stripboard jumper - 0 - +5 v>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                XDisplay1 = (OffsetX + x20first) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (OffsetX + x20last) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                'Call NewDrawBlackLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)

                'Call NewDrawBlackArcLineSegment(XDisplay1, YDisplay1, bmp)
                'Call NewDrawBlackArcLineSegment(XDisplay2, YDisplay2, bmp)

                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Black", bmp)

                Call DrawCircle(XDisplay1, YDisplay1, "Black", bmp)
                Call DrawCircle(XDisplay2, YDisplay2, "Black", bmp)

            ElseIf InStr(ListBox3.Items.Item(ii), "<stripboard jumper - 0 - +12 v>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                XDisplay1 = (OffsetX + x20first) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (OffsetX + x20last) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                'Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "Purple", bmp)

                'Call NewDrawPurpleArcLineSegment(XDisplay1, YDisplay1, bmp)
                'Call NewDrawPurpleArcLineSegment(XDisplay2, YDisplay2, bmp)

                Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "Purple", bmp)

                Call DrawCircle(XDisplay1, YDisplay1, "Purple", bmp)
                Call DrawCircle(XDisplay1, YDisplay2, "Purple", bmp)

            ElseIf InStr(ListBox3.Items.Item(ii), "<break>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                XDisplay1 = (OffsetX + x20first) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (OffsetX + x20last) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                'Call NewDrawIvoryLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)

                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Ivory", bmp)

            ElseIf InStr(ListBox3.Items.Item(ii), "<twixt break>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1)) '+ 25.4 / 2
                y20first = CSng(sp(3)) '+ 25.4 / 2
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1)) '+ 25.4 / 2
                y20last = CSng(sp(3)) '- 25.4 / 2

                XDisplay1 = (OffsetX + x20first) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (OffsetX + x20last) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                'Call NewDrawIvoryLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)

                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Ivory", bmp)

            End If
            sp = Split(ListBox3.Items.Item(ii), Chr(34))
            PenWidth = lastPenWidth
        Next
skip01:
        '********************************************************************
        '********************************************************************
        '********************************************************************
        '********************************************************************
        PictureBox1.Image = bmp
        graphics.Dispose()
        PictureBox1.Refresh()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        ToolFlag = "stripboard jumper - 0 - +5 v"
        PenWidth = CInt(TextBox23.Text)
        ListBox21.Items.Clear()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        ToolFlag = "stripboard jumper - gnd"
        PenWidth = CInt(TextBox23.Text)
        ListBox21.Items.Clear()
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        ToolFlag = "stripboard jumper - +5 v"
        PenWidth = CInt(TextBox23.Text)
        ListBox21.Items.Clear()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ToolFlag = "stripboard jumper - +12 v"
        PenWidth = CInt(TextBox23.Text)
        ListBox21.Items.Clear()
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ToolFlag = "stripboard jumper - 0 - +12 v"
        PenWidth = CInt(TextBox23.Text)
        ListBox21.Items.Clear()
    End Sub

    Private Sub DrawCircle(ByVal x1 As Single, ByVal y1 As Single, ByVal colour As String, ByVal bm As Bitmap)
        Dim graphics As Graphics = graphics.FromImage(bm)
        Dim tick_pen As New Pen(Color.Red, 3)
        Dim x11, y11 As Single

        Dim width As Single = 3
        Dim height As Single = 3

        Select Case colour
            Case "DarkGreen"
                tick_pen = New Pen(Color.DarkGreen, 3)
            Case "Orange"
                tick_pen = New Pen(Color.Orange, 3)
            Case "Purple"
                tick_pen = New Pen(Color.Purple, 3)
            Case "DarkRed"
                tick_pen = New Pen(Color.DarkRed, 3)
            Case "Red"
                tick_pen = New Pen(Color.Red, 3)
            Case "Black"
                tick_pen = New Pen(Color.Black, 3)
            Case "Ivory"
                tick_pen = New Pen(Color.Ivory, 3)
            Case "Blue"
                tick_pen = New Pen(Color.Blue, 3)
            Case "Gold"
                tick_pen = New Pen(Color.Gold, 3)
            Case "DarkGold"
                tick_pen = New Pen(Color.Goldenrod, 3)
            Case "Green"
                tick_pen = New Pen(Color.Green, 3)
            Case "HotPink"
                tick_pen = New Pen(Color.HotPink, 3)
        End Select

        x11 = x1 - width / 2
        y11 = y1 - height / 2

        graphics.DrawArc(tick_pen, x11, y11, width, height, 0, 360)
        'graphics.DrawArc(
    End Sub

    Private Sub DrawLine(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal colour As String, ByVal bm As Bitmap)
        Dim graphics As Graphics = graphics.FromImage(bm)
        Dim tick_pen As New Pen(Color.GreenYellow, PenWidth)

        Select Case colour
            Case "DarkGreen"
                tick_pen = New Pen(Color.DarkGreen, PenWidth)
            Case "Orange"
                tick_pen = New Pen(Color.Orange, PenWidth)
            Case "Purple"
                tick_pen = New Pen(Color.Purple, PenWidth)
            Case "DarkRed"
                tick_pen = New Pen(Color.DarkRed, PenWidth)
            Case "Red"
                tick_pen = New Pen(Color.Red, PenWidth)
            Case "Black"
                tick_pen = New Pen(Color.Black, PenWidth)
            Case "Ivory"
                tick_pen = New Pen(Color.Ivory, PenWidth)
            Case "Blue"
                tick_pen = New Pen(Color.Blue, PenWidth)
            Case "DarkGold"
                tick_pen = New Pen(Color.Goldenrod, PenWidth)
            Case "Gold"
                tick_pen = New Pen(Color.Gold, PenWidth)
            Case "Green"
                tick_pen = New Pen(Color.Green, PenWidth)
            Case "HotPink"
                tick_pen = New Pen(Color.HotPink, PenWidth)
        End Select

        Dim P1 As PointF
        P1.X = x1
        P1.Y = y1
        Dim P2 As PointF
        P2.X = x2
        P2.Y = y2
        graphics.DrawLine(tick_pen, P1, P2)

    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub FlipStripboardDisplay(ByVal ShowIndex As Integer)
        Dim bmp As New Bitmap(PictureBox1.Image)
        Dim graphics As Graphics = graphics.FromImage(bmp)

        graphics.Clear(Color.Ivory)

        PictureBox1.Image = bmp
        graphics.Dispose()
        PictureBox1.Refresh()

        'Exit Sub

        PenWidth = CSng(TextBox9.Text)


        Dim XDisplay1 As Single
        Dim YDisplay1 As Single

        Dim XDisplay2 As Single
        Dim YDisplay2 As Single
        Dim sp As Object

        Dim BoundariesFlag As Boolean = False

        Dim ii As Single
        Dim iii As Single
        'Dim j As Integer

        Dim x1, y1 As Integer

        Dim LoopCount As Integer = 0
        Dim TraceCount1 As Integer = 0
        Dim TraceCount2 As Integer = 0

        Dim EndsRho(3) As Double

        Dim RhoLimit As Double = 1 * 2 ^ 0.5

        '********************************************************************
        '********************************************************************

        'XBoardLimit = CInt(TextBox26.Text)
        'YBoardLimit = CInt(TextBox27.Text)
        'OffsetX = CInt(TextBox28.Text)
        'OffsetY = CInt(TextBox29.Text)
        DisplayScale = CSng(TextBox35.Text) '1 / CSng(TextBox30.Text)

        PinSpacing = 25.4 '* 10

        PenWidth = CSng(TextBox9.Text)

        '********************************************************************
        '********************************************************************

        Dim lastPenWidth As Single

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        'ListBox3.Items.Clear()
        ListBox4.Items.Clear()
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox19.Items.Clear()
        ListBox21.Items.Clear()

        Dim XOrigin, YOrigin As Double

        DisplayScale = CSng(TextBox35.Text) '1 / CSng(TextBox30.Text)



        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ComboBox7.Items.Clear()

        '********************************************************************
        '********************************************************************
        ''Dim InputPath As String = Mid(CurDir, 1, Len(CurDir) - 8) & "Stripboards\Footprints\"
        Dim InputPath As String = BasePath & "\Footprints\"

        Dim InputString As String

        '   load in the footprints

        FileOpen(100, InputPath & "Footprints 01.txt", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Do Until EOF(100)
            InputString = LineInput(100)
            'ListBox19.Items.Add(InputString)
            If InStr(InputString, "<footprint = ") > 0 Then
                sp = Split(InputString, Chr(34))
                ComboBox7.Items.Add(sp(1))
            End If
        Loop
        ComboBox7.Text = "component footprints"
        FileClose(100)
        '********************************************************************
        '********************************************************************
        FileClose()


        ListBox20.Items.Clear()

        XOrigin = OffsetX
        YOrigin = OffsetY

        x1 = OffsetX
        y1 = YBoardLimit - OffsetY

        XDisplay1 = OffsetX / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY + PinSpacing) / DisplayScale
        XDisplay2 = OffsetX / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        XDisplay1 = OffsetX / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale
        XDisplay2 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        XDisplay1 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY - YBoardLimit) / DisplayScale
        XDisplay2 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY + PinSpacing) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        XDisplay1 = (OffsetX + XBoardLimit) / DisplayScale
        YDisplay1 = (YBoardLimit + OffsetY + PinSpacing) / DisplayScale
        XDisplay2 = OffsetX / DisplayScale
        YDisplay2 = (YBoardLimit + OffsetY + PinSpacing) / DisplayScale

        Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        'PictureBox1.Image = bmp
        'graphics.Dispose()
        'PictureBox1.Refresh()
        'Exit Sub

        PenWidth = 9

        Dim ii1 As Integer = 0.75 * PinSpacing
        Dim ii2 As Integer = XBoardLimit - 0.75 * PinSpacing

        ''For ii = PinSpacing To XBoardLimit - PinSpacing Step XBoardLimit - 2 * PinSpacing
        'For iii = PinSpacing To YBoardLimit - PinSpacing Step PinSpacing
        For iii = 1 To YBoardLimit Step PinSpacing
            XDisplay1 = (OffsetX + (ii1 - 5)) / DisplayScale
            YDisplay1 = (YBoardLimit + OffsetY - (iii)) / DisplayScale
            XDisplay2 = (OffsetX + (ii2 + 5)) / DisplayScale
            YDisplay2 = (YBoardLimit + OffsetY - (iii)) / DisplayScale

            ''Call NewDrawDarkGoldLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)
            Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "DarkGold", bmp)

            ListBox3.Items.Add("<strip>")
            ListBox3.Items.Add(Chr(9) & "<penwidth =" & Chr(34) & Str(PenWidth) & Chr(34) & "/>")
            ListBox3.Items.Add(Chr(9) & "<begin x=" & Chr(34) & Str(ii1) & Chr(34) & " y=" & Chr(34) & Str(iii) & Chr(34) & "/>")
            ListBox3.Items.Add(Chr(9) & "<end   x=" & Chr(34) & Str(ii2) & Chr(34) & " y=" & Chr(34) & Str(iii) & Chr(34) & "/>")
            ListBox3.Items.Add("</strip>")


            XDisplay1 = (OffsetX + (ii)) / DisplayScale
            YDisplay1 = (YBoardLimit + OffsetY - (iii - 5)) / DisplayScale
            XDisplay2 = (OffsetX + (ii)) / DisplayScale
            YDisplay2 = (YBoardLimit + OffsetY - (iii + 5)) / DisplayScale

            'Call NewDrawGoldLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)


            XDisplay1 = (ii) / DisplayScale + OffsetX
            YDisplay1 = YBoardLimit + OffsetY - (iii) / DisplayScale
            ListBox20.Items.Add(Str(ii) & Chr(9) & Str(iii) & Chr(9) & Str(XDisplay1) & Chr(9) & Str(YDisplay1))

        Next
        ''Next

        PenWidth = 5

        'For ii = PinSpacing To XBoardLimit - PinSpacing Step PinSpacing
        'For ii = PinSpacing To XBoardLimit Step PinSpacing
        For ii = XBoardLimit - PinSpacing To 0 Step -PinSpacing

            'For iii = PinSpacing To YBoardLimit - PinSpacing Step PinSpacing
            For iii = 1 To YBoardLimit Step PinSpacing

                XDisplay1 = (OffsetX + (ii - 5)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - (iii)) / DisplayScale
                XDisplay2 = (OffsetX + (ii + 5)) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - (iii)) / DisplayScale

                'Call NewDrawGreenLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)

                XDisplay1 = (OffsetX + (ii)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - (iii)) / DisplayScale

                ''Call NewDrawBlackArcLineSegment(XDisplay1, YDisplay1, bmp)

                Call DrawCircle(XDisplay1, YDisplay1, "Black", bmp)

                XDisplay1 = (OffsetX + (ii)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - (iii - 5)) / DisplayScale
                XDisplay2 = (OffsetX + (ii)) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - (iii + 5)) / DisplayScale

                'Call NewDrawGreenLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)


                XDisplay1 = (ii) / DisplayScale + OffsetX
                YDisplay1 = YBoardLimit + OffsetY - (iii) / DisplayScale
                ListBox20.Items.Add(Str(ii) & Chr(9) & Str(iii) & Chr(9) & Str(XDisplay1) & Chr(9) & Str(YDisplay1))

            Next
        Next

        PenWidth = CSng(TextBox9.Text)


        'GoTo skip01
        '********************************************************************
        '********************************************************************
        Dim x2, y2 As Single
        Dim x20, y20 As Single
        Dim x20first, y20first As Single
        Dim x20last, y20last As Single
        '********************************************************************
        '********************************************************************

        'REDRAW THE DESIGN

        For ii = 0 To ListBox3.Items.Count - 1
            If InStr(ListBox3.Items.Item(ii), "<footprint = ") > 0 Then

                '   get the location

                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))

                x2 = CSng(sp(1))
                y2 = CSng(sp(3))

                '   find the footprint
                InputPath = BasePath & "\Footprints\"

                FileOpen(100, InputPath & "Footprints 01.txt", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
                Do Until EOF(100)
                    InputString = LineInput(100)
                    If InStr(InputString, ListBox3.Items.Item(ii)) > 0 Then
                        Do Until InStr(InputString, "</footprint>") > 0
                            InputString = LineInput(100)
                            If InStr(InputString, "<perimeter>") > 0 Then
                                'Dim PerimeterBegin As Boolean = False
                                Do Until InStr(InputString, "</perimeter>") > 0
                                    InputString = LineInput(100)
                                    If InStr(InputString, "</perimeter>") = 0 Then
                                        sp = Split(InputString, Chr(34))
                                        x20first = x2 + CSng(sp(1)) * PinSpacing ' / DisplayScale
                                        y20first = y2 + CSng(sp(3)) * PinSpacing ' / DisplayScale
                                        x20last = x2 + CSng(sp(5)) * PinSpacing ' / DisplayScale
                                        y20last = y2 + CSng(sp(7)) * PinSpacing ' / DisplayScale

                                        XDisplay1 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20first)) / DisplayScale
                                        YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                                        XDisplay2 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20last)) / DisplayScale
                                        YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                                        If ShowIndex < 2 Then
                                            PenWidth = 3
                                            ''Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Blue", bmp)
                                            PenWidth = lastPenWidth
                                        End If
                                    End If
                                Loop
                            ElseIf InStr(InputString, "<pins>") > 0 Then
                                Do Until InStr(InputString, "</pins>") > 0
                                    InputString = LineInput(100)
                                    If InStr(InputString, "</pins>") = 0 Then
                                        sp = Split(InputString, Chr(34))
                                        x20 = x2 + CSng(sp(3)) * PinSpacing '/ DisplayScale
                                        y20 = y2 + CSng(sp(5)) * PinSpacing '/ DisplayScale

                                        x20first = x20 - 10 '/ 2
                                        y20first = y20
                                        'sp = Split(ListBox21.Items.Item(iii), Chr(9))
                                        x20last = x20 + 10 '/ 2
                                        y20last = y20
                                        lastPenWidth = PenWidth
                                        PenWidth = 10 'CSng(TextBox24.Text)
                                        'XDisplay1 = (OffsetX + x20first) / DisplayScale
                                        'YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                                        'XDisplay2 = (OffsetX + x20last) / DisplayScale
                                        'YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                                        XDisplay1 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20first)) / DisplayScale
                                        YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                                        XDisplay2 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20last)) / DisplayScale
                                        YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                                        ''Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Gold", bmp)
                                        PenWidth = lastPenWidth

                                        XDisplay1 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20)) / DisplayScale
                                        YDisplay1 = (YBoardLimit + OffsetY - y20) / DisplayScale
                                        If ShowIndex < 2 Then
                                            ''Call NewDrawRedArcLineSegment(XDisplay1, YDisplay1, bmp)

                                            Call DrawCircle(XDisplay1, YDisplay1, "Red", bmp)
                                        End If
                                    End If
                                Loop
                            End If

                        Loop
                    End If
                Loop
                FileClose(100)
            ElseIf InStr(ListBox3.Items.Item(ii), "<stripboard jumper>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                'XDisplay1 = (OffsetX + x20first) / DisplayScale
                'YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                'XDisplay2 = (OffsetX + x20last) / DisplayScale
                'YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                XDisplay1 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20first)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20last)) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                ''Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "Orange", bmp)

                ''Call DrawCircle(XDisplay1, YDisplay1, "Orange", bmp)
                ''Call DrawCircle(XDisplay1, YDisplay2, "Orange", bmp)

            ElseIf InStr(ListBox3.Items.Item(ii), "<stripboard jumper - gnd>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                'XDisplay1 = (OffsetX + x20first) / DisplayScale
                'YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                'XDisplay2 = (OffsetX + x20last) / DisplayScale
                'YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                XDisplay1 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20first)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20last)) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                ''Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "DarkGreen", bmp)

                ''Call DrawCircle(XDisplay1, YDisplay1, "DarkGreen", bmp)
                ''Call DrawCircle(XDisplay1, YDisplay2, "DarkGreen", bmp)

            ElseIf InStr(ListBox3.Items.Item(ii), "<stripboard jumper - +5 v>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                'XDisplay1 = (OffsetX + x20first) / DisplayScale
                'YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                'XDisplay2 = (OffsetX + x20last) / DisplayScale
                'YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                XDisplay1 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20first)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20last)) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                ''Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "DarkRed", bmp)

                ''Call DrawCircle(XDisplay1, YDisplay1, "DarkRed", bmp)
                ''Call DrawCircle(XDisplay1, YDisplay2, "DarkRed", bmp)


            ElseIf InStr(ListBox3.Items.Item(ii), "<stripboard jumper - +12 v>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                'XDisplay1 = (OffsetX + x20first) / DisplayScale
                'YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                'XDisplay2 = (OffsetX + x20last) / DisplayScale
                'YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                XDisplay1 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20first)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20last)) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                ''Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "Red", bmp)

                ''Call DrawCircle(XDisplay1, YDisplay1, "Red", bmp)
                ''Call DrawCircle(XDisplay1, YDisplay2, "Red", bmp)

            ElseIf InStr(ListBox3.Items.Item(ii), "<stripboard jumper - 0 - +5 v>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                'XDisplay1 = (OffsetX + x20first) / DisplayScale
                'YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                'XDisplay2 = (OffsetX + x20last) / DisplayScale
                'YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                XDisplay1 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20first)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20last)) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                ''Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Black", bmp)

                ''Call DrawCircle(XDisplay1, YDisplay1, "Black", bmp)
                ''Call DrawCircle(XDisplay2, YDisplay2, "Black", bmp)

            ElseIf InStr(ListBox3.Items.Item(ii), "<stripboard jumper - 0 - +12 v>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                'XDisplay1 = (OffsetX + x20first) / DisplayScale
                'YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                'XDisplay2 = (OffsetX + x20last) / DisplayScale
                'YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                XDisplay1 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20first)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20last)) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                ''Call DrawLine(XDisplay1, YDisplay1, XDisplay1, YDisplay2, "Purple", bmp)

                ''Call DrawCircle(XDisplay1, YDisplay1, "Purple", bmp)
                ''Call DrawCircle(XDisplay1, YDisplay2, "Purple", bmp)

            ElseIf InStr(ListBox3.Items.Item(ii), "<break>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1))
                y20first = CSng(sp(3))
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1))
                y20last = CSng(sp(3))

                'XDisplay1 = (OffsetX + x20first) / DisplayScale
                'YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                'XDisplay2 = (OffsetX + x20last) / DisplayScale
                'YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                XDisplay1 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20first)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20last)) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Ivory", bmp)

            ElseIf InStr(ListBox3.Items.Item(ii), "<twixt break>") > 0 Then
                '   get the pen width
                lastPenWidth = PenWidth
                sp = Split(ListBox3.Items.Item(ii + 1), Chr(34))
                PenWidth = CSng(sp(1))
                sp = Split(ListBox3.Items.Item(ii + 2), Chr(34))
                x20first = CSng(sp(1)) '+ 25.4 / 2
                y20first = CSng(sp(3)) '+ 25.4 / 2
                sp = Split(ListBox3.Items.Item(ii + 3), Chr(34))
                x20last = CSng(sp(1)) '+ 25.4 / 2
                y20last = CSng(sp(3)) '- 25.4 / 2

                XDisplay1 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20first)) / DisplayScale
                YDisplay1 = (YBoardLimit + OffsetY - y20first) / DisplayScale
                XDisplay2 = (XBoardLimit + 4 * PinSpacing - (OffsetX + x20last)) / DisplayScale
                YDisplay2 = (YBoardLimit + OffsetY - y20last) / DisplayScale

                'Call NewDrawIvoryLineSegment(XDisplay1, YDisplay1, XDisplay2, YDisplay2, bmp)

                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Ivory", bmp)

            End If
            sp = Split(ListBox3.Items.Item(ii), Chr(34))
            PenWidth = lastPenWidth
        Next
skip01:
        '********************************************************************
        '********************************************************************
        '********************************************************************
        '********************************************************************
        PictureBox1.Image = bmp
        graphics.Dispose()
        PictureBox1.Refresh()
    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If Button4.Text = "Front" Then
            Call FlipStripboardDisplay(0)
            Button4.Text = "Back"
        Else
            Call RefreshStripboardDisplay(0)
            Button4.Text = "Front"
        End If

    End Sub


    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        ToolFlag = "twixt break"
        PenWidth = CInt(TextBox24.Text)
        ListBox21.Items.Clear()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '********************************************************************
        '********************************************************************
        '********************************************************************
        '********************************************************************


        Dim bmp As New Bitmap(PictureBox1.Image)
        Dim graphics As Graphics = graphics.FromImage(bmp)

        Dim CentreX As Integer = PictureBox1.Size.Width / 2
        Dim CentreY As Integer = PictureBox1.Size.Height / 2

        PenWidth = CSng(TextBox9.Text)

        Dim XDisplay1 As Single
        Dim YDisplay1 As Single

        Dim XDisplay2 As Single
        Dim YDisplay2 As Single

        Dim sp As Object

        Dim BoundariesFlag As Boolean = False

        Dim ii As Single
        Dim iii As Single
        'Dim j As Integer

        Dim x1, y1 As Integer ', x2, y2 As Integer

        'Dim rho11, rho12, rho21, rho22 As Double

        Dim LoopCount As Integer = 0
        Dim TraceCount1 As Integer = 0
        Dim TraceCount2 As Integer = 0

        'Dim rho As Double
        Dim EndsRho(3) As Double

        Dim RhoLimit As Double = 1 * 2 ^ 0.5

        '********************************************************************
        '********************************************************************

        'graphics.Clear(Color.Ivory)

        GroupBox4.Visible = True
        GroupBox5.Visible = True
        GroupBox3.Visible = True
        GroupBox7.Visible = True

        'Button4.Text = "Front"

        PinSpacing = 25.4 '* 10

        Dim xHoles As Integer = CInt(TextBox7.Text)
        Dim yHoles As Integer = CInt(TextBox8.Text)

        If xHoles < 25 Then xHoles = 25
        If yHoles < 5 Then yHoles = 5
        If xHoles > 85 Then xHoles = 85
        If yHoles > 60 Then yHoles = 60

        TextBox7.Text = Str(xHoles)
        TextBox8.Text = Str(yHoles)

        XBoardLimit = (xHoles + 1) * PinSpacing
        YBoardLimit = yHoles * PinSpacing

        TextBox26.Text = XBoardLimit
        TextBox27.Text = YBoardLimit

        TextBox26.Refresh()
        TextBox27.Refresh()

        OffsetX = 50 'CInt(TextBox28.Text)
        OffsetY = 50 'CInt(TextBox29.Text)
        DisplayScale = CSng(TextBox35.Text)

        PenWidth = 2

        'For ii = PinSpacing To XBoardLimit - PinSpacing Step PinSpacing
        ''For ii = PinSpacing To XBoardLimit - PinSpacing / 2 Step PinSpacing
        'For iii = PinSpacing To YBoardLimit - PinSpacing Step PinSpacing

        Dim ii1 As Integer = OffsetX - 15 '0.75 * PinSpacing
        Dim ii2 As Integer = OffsetX + XBoardLimit + 15 '- 0.75 * PinSpacing

        Dim xCount As Integer = 0
        For iii = 0 To XBoardLimit Step 5 * PinSpacing
            xCount = xCount + 1
            XDisplay1 = (OffsetX + (iii)) / DisplayScale
            YDisplay1 = (OffsetY - 15) / DisplayScale
            XDisplay2 = (OffsetX + (iii)) / DisplayScale
            YDisplay2 = (YBoardLimit + OffsetY + 35) / DisplayScale

            ''Call NewDrawBlackArcLineSegment(XDisplay1, YDisplay1, bmp)
            If xCount > 1 Then
                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "lightgreen", bmp)
            End If
        Next

        xCount = 0
        For iii = 0 To XBoardLimit Step 10 * PinSpacing
            xCount = xCount + 1
            XDisplay1 = (OffsetX + (iii)) / DisplayScale
            YDisplay1 = (OffsetY - 15) / DisplayScale
            XDisplay2 = (OffsetX + (iii)) / DisplayScale
            YDisplay2 = (YBoardLimit + OffsetY + 35) / DisplayScale

            ''Call NewDrawBlackArcLineSegment(XDisplay1, YDisplay1, bmp)
            If xCount > 1 Then
                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)
            End If
        Next

        ii1 = OffsetX - 15 '0.75 * PinSpacing
        ii2 = OffsetX + XBoardLimit + 15 '- 0.75 * PinSpacing

        Dim yCount As Integer = 0
        For iii = 0 To YBoardLimit Step 5 * PinSpacing
            yCount = yCount + 1
            YDisplay1 = (OffsetX + (iii)) / DisplayScale
            XDisplay1 = (OffsetX - 15) / DisplayScale
            YDisplay2 = (OffsetX + (iii)) / DisplayScale
            XDisplay2 = (XBoardLimit + OffsetX + 15) / DisplayScale

            ''Call NewDrawBlackArcLineSegment(XDisplay1, YDisplay1, bmp)
            If yCount > 1 Then
                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "lightgreen", bmp)
            End If
        Next

        yCount = 0
        For iii = 0 To YBoardLimit Step 10 * PinSpacing
            yCount = yCount + 1
            YDisplay1 = (OffsetX + (iii)) / DisplayScale
            XDisplay1 = (OffsetX - 15) / DisplayScale
            YDisplay2 = (OffsetX + (iii)) / DisplayScale
            XDisplay2 = (XBoardLimit + OffsetX + 15) / DisplayScale

            ''Call NewDrawBlackArcLineSegment(XDisplay1, YDisplay1, bmp)
            If yCount > 1 Then
                Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)
            End If
        Next

        GoTo skipallthis

        ''Dim ii1 As Integer = OffsetX - 15 '0.75 * PinSpacing
        ''Dim ii2 As Integer = OffsetX + XBoardLimit + 15 '- 0.75 * PinSpacing

        For iii = 4 * PinSpacing To YBoardLimit Step 5 * PinSpacing
            XDisplay1 = ii1 / DisplayScale
            YDisplay1 = (YBoardLimit + OffsetY - (iii)) / DisplayScale
            XDisplay2 = ii2 / DisplayScale
            YDisplay2 = (YBoardLimit + OffsetY - (iii)) / DisplayScale

            Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        Next

        For iii = 4 * PinSpacing To YBoardLimit Step 10 * PinSpacing
            XDisplay1 = ii1 / DisplayScale
            YDisplay1 = (YBoardLimit + OffsetY - (iii)) / DisplayScale
            XDisplay2 = ii2 / DisplayScale
            YDisplay2 = (YBoardLimit + OffsetY - (iii)) / DisplayScale

            Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "light green", bmp)

        Next

        ii1 = OffsetY - 15 '0.75 * PinSpacing
        ii2 = OffsetY + YBoardLimit + 35 '- 0.75 * PinSpacing

        For iii = 4 * PinSpacing To XBoardLimit Step 5 * PinSpacing
            YDisplay1 = ii1 / DisplayScale
            XDisplay1 = (XBoardLimit + OffsetX - (iii)) / DisplayScale
            YDisplay2 = ii2 / DisplayScale
            XDisplay2 = (XBoardLimit + OffsetX - (iii)) / DisplayScale

            Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "light green", bmp)

        Next

        For iii = 4 * PinSpacing To XBoardLimit Step 10 * PinSpacing
            YDisplay1 = ii1 / DisplayScale
            XDisplay1 = (XBoardLimit + OffsetX - (iii)) / DisplayScale
            YDisplay2 = ii2 / DisplayScale
            XDisplay2 = (XBoardLimit + OffsetX - (iii)) / DisplayScale

            Call DrawLine(XDisplay1, YDisplay1, XDisplay2, YDisplay2, "Red", bmp)

        Next
skipallthis:
        PenWidth = CSng(TextBox9.Text)

        PictureBox1.Image = bmp
        graphics.Dispose()
        PictureBox1.Refresh()

        Exit Sub

        '********************************************************************
    End Sub
End Class
