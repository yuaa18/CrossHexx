Imports System.Diagnostics
Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices


Public Class Form1
    Private WithEvents MouseHook As New MouseHookClass

    '----------オーバレイ宣言----------------
    Public Const HWND_TOPMOST = (-1)
    Public Const SWP_NOSIZE = &H1&
    Public Const SWP_NOMOVE = &H2&

    Public Const GWL_EXSTYLE As Long = (-20)

    Public Const WS_EX_LAYERED = &H80000
    Public Const WS_EX_TRANSPARENT = &H20
    Private Const LWA_COLORKEY = &H1
    Public Const LWA_ALPHA = &H2

    Const TOPMOST_FLAGS As UInteger = (SWP_NOSIZE Or SWP_NOMOVE)

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetWindowPos(ByVal hWnd As IntPtr,
        ByVal hWndInsertAfter As IntPtr,
        ByVal X As Integer,
        ByVal Y As Integer,
        ByVal cx As Integer,
        ByVal cy As Integer,
        ByVal uFlags As UInteger) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetWindowLong(hWnd As IntPtr,
        <MarshalAs(UnmanagedType.I4)> nIndex As Integer,
        dwNewLong As IntPtr) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowLong(hWnd As IntPtr,
        <MarshalAs(UnmanagedType.I4)> nIndex As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetLayeredWindowAttributes(hwnd As IntPtr, crKey As UInteger, bAlpha As Byte, dwFlags As UInteger) As Boolean
    End Function
    '-------------ここまで-------------------

    '-------------ウィンドウ取得------------

    Public Declare Function FindWindowA Lib "user32" (ByVal cnm As String, ByVal cap As String) As Integer

    Dim stat As Integer
    Dim count As Integer

    Dim Rect2 As Rectangle
    Dim DisWidth As Integer
    Dim DisHeight As Integer
    Dim FirstWidth As Integer
    Dim FirstHeight As Integer
    Dim NumWidth As Integer
    Dim NumHeight As Integer

    Dim DisWidth2 As Integer
    Dim DisHeight2 As Integer

    Dim colorR As Integer
    Dim colorG As Integer
    Dim colorB As Integer

    'クロスヘアのサイズと種類の変数を用意
    Dim c_size As String
    Dim c_type As String

    '1 = 画像
    Dim statImage As Integer

    Dim fWait As Integer
    Dim rWait As Integer

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function GetWindowRect(ByVal hWnd As IntPtr,
                                          ByRef lpRect As RECT) _
                                          As Boolean
    End Function

    Private Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure
    '---------------------------------------

    'プレビュー用をfとする
    Dim f As New Form2()

    Dim setting1(,) As Integer = New Integer(4, 3) {
                                          {1, My.Settings.type1, My.Settings.size1, My.Settings.color1},
                                          {2, My.Settings.type2, My.Settings.size2, My.Settings.color2},
                                          {3, My.Settings.type3, My.Settings.size3, My.Settings.color3},
                                          {4, My.Settings.type4, My.Settings.size4, My.Settings.color4},
                                          {5, My.Settings.type5, My.Settings.size5, My.Settings.color5}}

    'iniの読み取り
    Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (
    ByVal lpApplicationName As String,
    ByVal lpKeyName As String,
    ByVal nDefault As Integer,
    ByVal lpFileName As String) As Integer


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'デスクトップのサイズ取得
        Rect2 = Screen.PrimaryScreen.Bounds
        'ポジションに設定
        DisHeight = Rect2.Height / 2 - (Form2.Height / 2)
        DisWidth = Rect2.Width / 2 - (Form2.Width / 2)

        statImage = 0
        rWait = 0

        FirstHeight = DisHeight
        FirstWidth = DisWidth


        colorR = Panel1.BackColor.R
        colorG = Panel1.BackColor.G
        colorB = Panel1.BackColor.B

        TextBox1.Text = colorR
        TextBox2.Text = colorG
        TextBox3.Text = colorB

        Dim maxHeight As Integer = Rect2.Height / 2
        Dim maxWidth As Integer = Rect2.Width / 2

        '=======x軸y軸=======
        NumericUpDown1.Maximum = maxHeight
        NumericUpDown1.Minimum = -maxHeight
        NumericUpDown2.Maximum = maxWidth
        NumericUpDown2.Minimum = -maxWidth

        '=======オフセット=======
        NumericUpDown5.Maximum = maxHeight
        NumericUpDown5.Minimum = -maxHeight
        NumericUpDown6.Maximum = maxWidth
        NumericUpDown6.Minimum = -maxWidth

        '=======save/load=======
        NumericUpDown7.Maximum = maxHeight
        NumericUpDown7.Minimum = -maxHeight
        NumericUpDown8.Maximum = maxWidth
        NumericUpDown8.Minimum = -maxWidth


        NumericUpDown1.Value = My.Settings.y
        NumericUpDown2.Value = My.Settings.x

        NumericUpDown5.Value = My.Settings.y_offset
        NumericUpDown6.Value = My.Settings.x_offset


        ldsetName()

        SetWindowLong(Form2.Handle, GWL_EXSTYLE, GetWindowLong(Form2.Handle, GWL_EXSTYLE) Xor WS_EX_LAYERED Xor WS_EX_TRANSPARENT)

        SetLayeredWindowAttributes(Form2.Handle, 0, 255, LWA_ALPHA)
        SetWindowPos(Form2.Handle, HWND_TOPMOST, 0, 0, 0, 0,
            TOPMOST_FLAGS)


    End Sub



    Private Sub ldsetName()


        'Addで一つ一つ追加
        'BeginUpdateを使用
        ListBox1.Items.Clear()
        '再描画しないようにする
        ListBox1.BeginUpdate()
        '配列の内容を一つ一つ追加する
        ListBox1.Items.Add(My.Settings.name1)
        ListBox1.Items.Add(My.Settings.name2)
        ListBox1.Items.Add(My.Settings.name3)
        ListBox1.Items.Add(My.Settings.name4)
        ListBox1.Items.Add(My.Settings.name5)
        '再描画するようにする
        ListBox1.EndUpdate()

    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim cd As New ColorDialog()
        If cd.ShowDialog() = DialogResult.OK Then

            '選択された色の取得

            TextBox1.Text = cd.Color.R
            TextBox2.Text = cd.Color.G
            TextBox3.Text = cd.Color.B
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Button1.Enabled = False
        Button2.Enabled = True
        '起動中ならば1
        stat = 1

        x_axis()
        y_axis()

        If RadioButton1.Checked = True Then
            'クロスヘア
            c_type = "cross"
            If RadioButton4.Checked = True Then
                c_size = "small"
            Else
                c_size = "large"
            End If
            crosshair(c_size, c_type)
        ElseIf RadioButton2.Checked = True Then
            'ドット
            c_type = "dot"
            If RadioButton4.Checked = True Then
                c_size = "small"
            Else
                c_size = "large"
            End If
            crosshair(c_size, c_type)
        ElseIf RadioButton3.Checked = True Then
            'ドット&サークル
            c_type = "circle"
            If RadioButton4.Checked = True Then
                c_size = "small"
            Else
                c_size = "large"
            End If
            crosshair(c_size, c_type)
        ElseIf RadioButton6.Checked = True Then

            showimage()

        End If


        Form2.Show()
        Form2.Location = New Point(DisWidth, DisHeight)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Button1.Enabled = True
        Button2.Enabled = False
        '停止中ならば0
        stat = 0
        rWait = 0
        If statImage = 1 Then
            Form2.Close()
        Else
            Form2.Hide()
        End If

    End Sub


    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If statImage = 1 Then
            setWin()
        End If

        statImage = 0
        If stat = 1 Then
            '起動中ならば表示
            c_type = "cross"
            If RadioButton4.Checked = True Then
                c_size = "small"
            ElseIf RadioButton4.Checked = False Then
                c_size = "large"
            End If
            crosshair(c_size, c_type)

            Form2.Show()
            Form2.Location = New Point(DisWidth, DisHeight)

        End If

        'プレビュー
        If TabControl1.SelectedIndex = 0 Then
            If RadioButton4.Checked = True Then
                preview1()
            Else
                bigpreview1()
            End If
        End If

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If statImage = 1 Then
            setWin()
        End If
        statImage = 0
        If stat = 1 Then
            c_type = "dot"
            If RadioButton4.Checked = True Then
                c_size = "small"
            ElseIf RadioButton4.Checked = False Then
                c_size = "large"
            End If
            crosshair(c_size, c_type)

            Form2.Show()
            Form2.Location = New Point(DisWidth, DisHeight)

        End If

        If TabControl1.SelectedIndex = 0 Then
            If RadioButton4.Checked = True Then
                preview2()
            Else
                bigpreview2()
            End If
        End If

    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If statImage = 1 Then
            setWin()
        End If

        statImage = 0

        If stat = 1 Then
            c_type = "circle"
            If RadioButton4.Checked = True Then
                c_size = "small"
            ElseIf RadioButton4.Checked = False Then
                c_size = "large"
            End If
            crosshair(c_size, c_type)

            Form2.Show()
            Form2.Location = New Point(DisWidth, DisHeight)

        End If


        If RadioButton4.Checked = True Then
            preview3()
        Else
            bigpreview3()
        End If



    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged
        '画像の描写
        statImage = 1

        If stat = 1 Then

            Form2.Close()
            showimage()
            Form2.Show()
            Form2.Location = New Point(DisWidth, DisHeight)

        End If

        If RadioButton6.Checked = False Then
            Form2.PictureBox1.Visible = False
        End If

    End Sub

    Private Sub crosshair(ByVal sizes_in As String, ByVal types_in As String)

        If CheckBox2.Checked Then
            fWait = TextBox6.Text
            System.Threading.Thread.Sleep(fWait * 1000)
        End If

        rWait = 1

        Dim points() As Point = {}


        'GraphicsPathの作成
        Dim path As New GraphicsPath
        path.StartFigure()

        If types_in = "cross" Then 'クロスヘア

            If sizes_in = "small" Then

                points =
                {New Point(17, 16),
                New Point(2, 17),
                New Point(17, 18),
                New Point(17, 32),
                New Point(18, 18),
                New Point(32, 17),
                New Point(18, 16),
                New Point(17, 2)}

                path.AddLines(points)


            ElseIf sizes_in = "large" Then
                points =
                {New Point(16, 16),
                New Point(2, 16),
                New Point(2, 19),
                New Point(16, 19),
                New Point(16, 32),
                New Point(19, 32),
                New Point(19, 19),
                New Point(32, 19),
                New Point(32, 16),
                New Point(19, 16),
                New Point(19, 2),
                New Point(16, 2)}

                path.AddLines(points)


            End If

        ElseIf types_in = "dot" Then
            If sizes_in = "small" Then
                path.AddEllipse(New Rectangle(15, 15, 4, 4))
            ElseIf sizes_in = "large" Then
                path.AddEllipse(New Rectangle(13, 13, 8, 8))
            End If

        ElseIf types_in = "circle" Then
            If sizes_in = "small" Then
                path.AddEllipse(New Rectangle(2, 2, 30, 30))
                path.AddEllipse(New Rectangle(3, 3, 28, 28))
                path.AddEllipse(New Rectangle(15, 15, 4, 4))
            ElseIf sizes_in = "large" Then
                path.AddEllipse(New Rectangle(0, 0, 34, 34))
                path.AddEllipse(New Rectangle(2, 2, 30, 30))
                path.AddEllipse(New Rectangle(13, 13, 8, 8))
            End If
        End If


        Form2.Region = New Region(path)
    End Sub

    Private Sub showimage()

        setWin()


        Dim points() As Point = {}


        'GraphicsPathの作成
        Dim path As New GraphicsPath
        path.StartFigure()

        points =
        {New Point(0, 0),
        New Point(0, h),
        New Point(w, h),
        New Point(w, 0)}

        path.AddLines(points)


        Form2.Region = New Region(path)
        Form2.PictureBox1.Image = PictureBox2.Image


        If w > 0 And h > 0 Then
            Form2.Size = New Size(w, h)
            Form2.PictureBox1.Width = w
            Form2.PictureBox1.Height = h
        End If



    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 0 Then
            If RadioButton4.Checked Then
                If RadioButton1.Checked = True Then
                    preview1()
                ElseIf RadioButton2.Checked = True Then
                    preview2()
                ElseIf RadioButton3.Checked = True Then
                    preview3()
                End If
            Else
                If RadioButton1.Checked = True Then
                    preview1()
                ElseIf RadioButton2.Checked = True Then
                    bigpreview2()
                ElseIf RadioButton3.Checked = True Then
                    bigpreview3()
                End If
            End If
        Else
            f.Hide()
        End If
    End Sub

    Private Sub preview1()
        '---------クロスヘアプレビュー用---------
        Dim points() As Point =
            {New Point(17, 16),
             New Point(2, 17),
             New Point(17, 18),
             New Point(17, 32),
             New Point(18, 18),
             New Point(32, 17),
             New Point(18, 16),
             New Point(17, 2)}

        Dim types() As Byte =
            {Drawing.Drawing2D.PathPointType.Line,
            Drawing.Drawing2D.PathPointType.Line,
            Drawing.Drawing2D.PathPointType.Line,
             Drawing.Drawing2D.PathPointType.Line,
            Drawing.Drawing2D.PathPointType.Line,
            Drawing.Drawing2D.PathPointType.Line,
             Drawing.Drawing2D.PathPointType.Line,
            Drawing.Drawing2D.PathPointType.Line}
        'GraphicsPathの作成
        Dim path As New Drawing2D.GraphicsPath(points, types)
        f.Region = New Region(path)
        'TopLevelをFalseにする
        f.TopLevel = False
        'フォームのコントロールに追加する
        Me.Controls.Add(f)
        'フォームを表示する
        If TabControl1.SelectedIndex = 0 Then
            f.Show()
        End If
        f.Location = New Point(150, 216)
        '最前面へ移動
        f.BringToFront()
    End Sub

    Private Sub preview2()
        SetWindowLong(Form2.Handle, GWL_EXSTYLE, GetWindowLong(Form2.Handle, GWL_EXSTYLE) Xor WS_EX_LAYERED Xor WS_EX_TRANSPARENT)

        SetLayeredWindowAttributes(Form2.Handle, 0, 255, LWA_ALPHA)
        SetWindowPos(Form2.Handle, HWND_TOPMOST, 0, 0, 0, 0,
            TOPMOST_FLAGS)
        Dim path As New System.Drawing.Drawing2D.GraphicsPath()
        '丸を描く
        path.AddEllipse(New Rectangle(15, 15, 4, 4))

        f.Region = New Region(path)
        'TopLevelをFalseにする
        f.TopLevel = False
        'フォームのコントロールに追加する
        Me.Controls.Add(f)
        'フォームを表示する
        If TabControl1.SelectedIndex = 0 Then
            f.Show()
        End If
        f.Location = New Point(150, 216)
        '最前面へ移動
        f.BringToFront()

    End Sub

    Private Sub preview3()
        SetWindowLong(Form2.Handle, GWL_EXSTYLE, GetWindowLong(Form2.Handle, GWL_EXSTYLE) Xor WS_EX_LAYERED Xor WS_EX_TRANSPARENT)

        SetLayeredWindowAttributes(Form2.Handle, 0, 255, LWA_ALPHA)
        SetWindowPos(Form2.Handle, HWND_TOPMOST, 0, 0, 0, 0,
            TOPMOST_FLAGS)
        Dim path As New System.Drawing.Drawing2D.GraphicsPath()
        '丸を描く
        path.AddEllipse(New Rectangle(2, 2, 30, 30))
        path.AddEllipse(New Rectangle(3, 3, 28, 28))
        path.AddEllipse(New Rectangle(15, 15, 4, 4))

        f.Region = New Region(path)
        'TopLevelをFalseにする
        f.TopLevel = False
        'フォームのコントロールに追加する
        Me.Controls.Add(f)
        'フォームを表示する
        If TabControl1.SelectedIndex = 0 Then
            f.Show()
        End If
        f.Location = New Point(150, 216)
        '最前面へ移動
        f.BringToFront()
    End Sub

    Private Sub bigpreview1()

        Dim points() As Point =
        {New Point(16, 16),
         New Point(2, 16),
         New Point(2, 19),
         New Point(16, 19),
         New Point(16, 32),
         New Point(19, 32),
         New Point(19, 19),
         New Point(32, 19),
         New Point(32, 16),
         New Point(19, 16),
          New Point(19, 2),
         New Point(16, 2)}

        Dim types() As Byte =
            {Drawing.Drawing2D.PathPointType.Line,
            Drawing.Drawing2D.PathPointType.Line,
            Drawing.Drawing2D.PathPointType.Line,
             Drawing.Drawing2D.PathPointType.Line,
             Drawing.Drawing2D.PathPointType.Line,
             Drawing.Drawing2D.PathPointType.Line,
             Drawing.Drawing2D.PathPointType.Line,
            Drawing.Drawing2D.PathPointType.Line,
            Drawing.Drawing2D.PathPointType.Line,
             Drawing.Drawing2D.PathPointType.Line,
             Drawing.Drawing2D.PathPointType.Line,
            Drawing.Drawing2D.PathPointType.Line}
        'GraphicsPathの作成
        Dim path As New Drawing2D.GraphicsPath(points, types)

        f.Region = New Region(path)
        'TopLevelをFalseにする
        f.TopLevel = False
        'フォームのコントロールに追加する
        Me.Controls.Add(f)
        'フォームを表示する
        If TabControl1.SelectedIndex = 0 Then
            f.Show()
        End If
        f.Location = New Point(150, 216)
        '最前面へ移動
        f.BringToFront()
    End Sub

    Private Sub bigpreview2()

        Dim path As New System.Drawing.Drawing2D.GraphicsPath()
        '丸を描く
        path.AddEllipse(New Rectangle(13, 13, 8, 8))

        f.Region = New Region(path)
        'TopLevelをFalseにする
        f.TopLevel = False
        'フォームのコントロールに追加する
        Me.Controls.Add(f)
        'フォームを表示する
        If TabControl1.SelectedIndex = 0 Then
            f.Show()
        End If
        f.Location = New Point(150, 216)
        '最前面へ移動
        f.BringToFront()
    End Sub

    Private Sub bigpreview3()

        Dim path As New System.Drawing.Drawing2D.GraphicsPath()
        '丸を描く
        path.AddEllipse(New Rectangle(0, 0, 34, 34))
        path.AddEllipse(New Rectangle(2, 2, 30, 30))
        path.AddEllipse(New Rectangle(13, 13, 8, 8))

        f.Region = New Region(path)
        'TopLevelをFalseにする
        f.TopLevel = False
        'フォームのコントロールに追加する
        Me.Controls.Add(f)
        'フォームを表示する
        If TabControl1.SelectedIndex = 0 Then
            f.Show()
        End If
        f.Location = New Point(150, 216)
        '最前面へ移動
        f.BringToFront()
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        'Medium

        'Form2.Hide()
        If RadioButton1.Checked = True Then
            preview1()
        ElseIf RadioButton2.Checked = True Then
            preview2()
        ElseIf RadioButton3.Checked = True Then
            preview3()
        End If
        If stat = 1 And statImage = 0 Then
            If RadioButton4.Checked = True Then
                c_size = "small"
                If RadioButton1.Checked = True Then
                    c_type = "cross"
                ElseIf RadioButton2.Checked = True Then
                    c_type = "dot"
                ElseIf RadioButton3.Checked = True Then
                    c_type = "circle"
                End If
                crosshair(c_size, c_type)

                Form2.Show()
                Form2.Location = New Point(DisWidth, DisHeight)
            End If
        End If
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        'large

        'Form2.Hide()
        If RadioButton1.Checked = True Then
            bigpreview1()
        ElseIf RadioButton2.Checked = True Then
            bigpreview2()
        ElseIf RadioButton3.Checked = True Then
            bigpreview3()
        End If
        If stat = 1 And statImage = 0 Then
            If RadioButton5.Checked = True Then
                c_size = "large"
                If RadioButton1.Checked = True Then
                    c_type = "cross"
                ElseIf RadioButton2.Checked = True Then
                    c_type = "dot"
                ElseIf RadioButton3.Checked = True Then
                    c_type = "circle"
                End If
                crosshair(c_size, c_type)

                Form2.Show()
                Form2.Location = New Point(DisWidth, DisHeight)
            End If
        End If
    End Sub



    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        If TextBox1.Text > 256 Then
            TextBox1.Text = 255
        End If
        colorR = TextBox1.Text
        colorset()
    End Sub


    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text > 256 Then
            TextBox2.Text = 255
        End If
        colorG = TextBox2.Text
        colorset()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text > 256 Then
            TextBox3.Text = 255
        End If
        colorB = TextBox3.Text
        colorset()
    End Sub

    Private Sub colorset()
        Form2.BackColor = Color.FromArgb(colorR, colorG, colorB)
        Panel1.BackColor = Color.FromArgb(colorR, colorG, colorB)
        f.BackColor = Color.FromArgb(colorR, colorG, colorB)
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged

        y_axis()


    End Sub


    Private Sub NumericUpDown5_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown5.ValueChanged

        y_axis()

    End Sub

    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
        x_axis()
    End Sub

    Private Sub NumericUpDown6_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown6.ValueChanged
        x_axis()
    End Sub

    Public Sub y_axis()
        NumHeight = FirstHeight

        If (NumericUpDown1.Value + NumericUpDown5.Value) > (Integer.Parse(NumericUpDown1.Text) + Integer.Parse(NumericUpDown5.Text)) Then
            '上へ移動
            Form2.Top -= 1
            DisHeight -= 1
        ElseIf (NumericUpDown1.Value + NumericUpDown5.Value) < (Integer.Parse(NumericUpDown1.Text) + Integer.Parse(NumericUpDown5.Text)) Then
            Form2.Top += 1
            DisHeight += 1
        End If
        NumHeight -= (NumericUpDown1.Value + NumericUpDown5.Value)
        DisHeight = FirstHeight - (NumericUpDown1.Value + NumericUpDown5.Value)
        Form2.Location = New Point(DisWidth, NumHeight)

    End Sub

    Public Sub x_axis()
        NumWidth = FirstWidth

        If (NumericUpDown2.Value + NumericUpDown6.Value) > (Integer.Parse(NumericUpDown2.Text) + Integer.Parse(NumericUpDown6.Text)) Then
            Form2.Left += 1
            DisWidth += 1
        ElseIf (NumericUpDown2.Value + NumericUpDown6.Value) < (Integer.Parse(NumericUpDown2.Text) + Integer.Parse(NumericUpDown6.Text)) Then
            Form2.Left -= 1
            DisWidth -= 1
        End If
        NumWidth += (NumericUpDown2.Value + NumericUpDown6.Value)
        DisWidth = FirstWidth + (NumericUpDown2.Value + NumericUpDown6.Value)
        Form2.Location = New Point(NumWidth, DisHeight)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        NumericUpDown1.Value = 0
        NumericUpDown2.Value = 0
        DisWidth = FirstWidth
        DisHeight = FirstHeight
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click


        NumericUpDown5.Value = 0
        NumericUpDown6.Value = 0


    End Sub


    Private Sub ComboBox1_Click(sender As Object, e As EventArgs) Handles ComboBox1.Click
        If count = 1 Then
            ComboBox1.Items.Clear()
            count = 0
        End If
        If count = 0 Then
            For Each item As Process In Process.GetProcesses()
                If item.MainWindowHandle <> IntPtr.Zero Then
                    ComboBox1.Items.Add(item.MainWindowTitle)
                    ComboBox1.Items.Remove("")
                    count = 1
                End If
            Next
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim hwindow As Integer
        Dim x = ComboBox1.Text
        hwindow = FindWindowA(vbNullString, x)
        Dim nRect As RECT

        Call GetWindowRect(hwindow, nRect)
        If nRect.Top = -32000 Then
            MsgBox("最小化されています")
        ElseIf nRect.Top = 0 And nRect.Bottom = 0 And nRect.Right = 0 And nRect.Left = 0 Then
            MsgBox("取得失敗")
        ElseIf x = "" Then
        Else
            DisHeight = (nRect.Bottom + SystemInformation.CaptionHeight - nRect.Top) / 2 + nRect.Top - 17
            DisWidth = (nRect.Right - nRect.Left) / 2 + nRect.Left - 17
            NumericUpDown1.Value = FirstHeight - DisHeight
            NumericUpDown2.Value = DisWidth - FirstWidth
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        'リンク先に移動したことにする
        LinkLabel1.LinkVisited = True
        'ブラウザで開く
        System.Diagnostics.Process.Start("http://mjh.blog.jp/")
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Me.TopMost = True
        Else
            Me.TopMost = False
        End If

    End Sub

    '---------設定class
    Public Class Settings
        Private _text As String
        Private _number As Integer

        Public Property Text() As String
            Get
                Return _text
            End Get
            Set(ByVal Value As String)
                _text = Value
            End Set
        End Property

        Public Property Number() As Integer
            Get
                Return _number
            End Get
            Set(ByVal Value As Integer)
                _number = Value
            End Set
        End Property

        Public Sub New()
            _text = "Text"
            _number = 0
        End Sub
    End Class
    '-----------------------------

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        My.Settings.y = NumericUpDown1.Value
        My.Settings.x = NumericUpDown2.Value
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        My.Settings.y_offset = NumericUpDown5.Value
        My.Settings.x_offset = NumericUpDown6.Value
    End Sub


    Dim w As Integer = 0
    Dim h As Integer = 0

    Dim oriw As Integer
    Dim orih As Integer
    Dim ofd As New OpenFileDialog()
    Dim inif As New OpenFileDialog()
    Dim img As Image

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        '[ファイルの種類]に表示される選択肢を指定する
        '指定しないとすべてのファイルが表示される
        ofd.Filter =
            "イメージファイル(*.gif;*.jpg;*.jpeg;*.bmp;*.wmf*.png)|*.gif;*.jpg;*.jpeg;*.bmp;*.wmf;*.png|すべてのファイル(*.*)|*.*"
        '[ファイルの種類]ではじめに
        '「すべてのファイル」が選択されているようにする
        ofd.FilterIndex = 1
        'タイトルを設定する
        ofd.Title = "開くファイルを選択してください"
        'ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
        ofd.RestoreDirectory = True



        'ダイアログを表示する
        If ofd.ShowDialog() = DialogResult.OK Then
            showpre()
            Try
                'NumericUpDownに画像の大きさを設定する
                NumericUpDown3.Value = oriw
                NumericUpDown4.Value = orih

            Catch ex As ArgumentOutOfRangeException

                MsgBox("サイズが大きすぎます 10000×10000以下で選択してください", MsgBoxStyle.Exclamation)

            End Try

        End If
    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged
        If (RadioButton7.Checked = True) And Not (PictureBox2.Image Is Nothing) Then
            showpre()
        End If
    End Sub

    Private Sub RadioButton8_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton8.CheckedChanged
        If (RadioButton8.Checked = True) And Not (PictureBox2.Image Is Nothing) Then
            showpre()
        End If
    End Sub

    Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged
        If (RadioButton9.Checked = True) And Not (PictureBox2.Image Is Nothing) Then
            showpre()
        End If
    End Sub
    Private Function showpre()

        '画像ファイルを読み込んで、Imageオブジェクトとして取得する
        Try
            img = Image.FromFile(ofd.FileName)
        Catch ex As System.OutOfMemoryException
            MsgBox("ファイルが正しくありません。", MsgBoxStyle.Exclamation)
            Return Nothing
        Catch ex As System.ArgumentException
            MsgBox("ファイルの形式が無効です。", MsgBoxStyle.Exclamation)
            Return Nothing
        Catch ex As System.IO.FileNotFoundException
            MsgBox("ファイルが存在しません。", MsgBoxStyle.Exclamation)
            Return Nothing
        End Try

        'ファイルパスをtextbox5に入れる
        TextBox5.Text = ofd.FileName

        '変数oriwとorihに画像のオリジナルの大きさを入れておく
        oriw = img.Width
        orih = img.Height
        If RadioButton7.Checked = True Then
            'PictureBox2の大きさに合わせる計算
            If (148 < img.Width Or 148 < img.Height) Then
                PictureBox2.Width = 148
                PictureBox2.Height = 148
                If (img.Width >= img.Height) Then
                    w = img.Width / (img.Width / 148)
                    h = img.Height / (img.Width / 148)
                Else
                    w = img.Width / (img.Height / 148)
                    h = img.Height / (img.Height / 148)

                End If
            Else
                w = img.Width
                h = img.Height
            End If

        ElseIf RadioButton8.Checked = True Then
            w = img.Width
            h = img.Height

        ElseIf RadioButton9.Checked = True Then
            w = NumericUpDown3.Value
            h = NumericUpDown4.Value
        End If

        Dim canvas As New Bitmap(w, h)
        'ImageオブジェクトのGraphicsオブジェクトを作成する
        Dim g As Graphics = Graphics.FromImage(canvas)

        '画像をcanvasの座標(0, 0)の位置に描画する
        g.DrawImage(img, 0, 0, w, h)
        'PictureBox2に表示する
        PictureBox2.Image = canvas

        'Imageオブジェクトのリソースを解放する
        img.Dispose()

        'Graphicsオブジェクトのリソースを解放する
        g.Dispose()


        If stat = 1 And RadioButton6.Checked = True Then

            showimage()
            Form2.Close()
            showimage()
            Form2.Show()
            Form2.Location = New Point(DisWidth, DisHeight)
            'If RadioButton6.Checked = False Then
            '    Form2.Hide()
            'End If
        End If
    End Function

    Private Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown3.ValueChanged
        If Not (PictureBox2.Image Is Nothing) Then
            showpre()
        End If
    End Sub

    Private Sub NumericUpDown4_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown4.ValueChanged
        If Not (PictureBox2.Image Is Nothing) Then
            showpre()
        End If
    End Sub


    Private Sub TabPage3_DragEnter(sender As Object, e As DragEventArgs) Handles TabPage3.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            'ドラッグされたデータ形式を調べ、ファイルのときはコピーとする
            e.Effect = DragDropEffects.Copy
        Else
            'ファイル以外は受け付けない
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub TabPage3_DragDrop(sender As Object, e As DragEventArgs) Handles TabPage3.DragDrop
        'コントロール内にドロップされたとき実行される
        'ドロップされたすべてのファイル名を取得する
        Dim fileName As String() = CType(
            e.Data.GetData(DataFormats.FileDrop, False),
            String())

        '配列をStringに変換
        Dim fN As String = fileName(0)
        ofd.FileName = fN
        showpre()

        Try
            'NumericUpDownに画像の大きさを設定する
            NumericUpDown3.Value = oriw
            NumericUpDown4.Value = orih

        Catch ex As ArgumentOutOfRangeException

            MsgBox("サイズが大きすぎます 10000×10000以下で選択してください", MsgBoxStyle.Exclamation)

        End Try

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        Dim wType As Integer
        Dim wSize As Integer
        Dim wColor As String
        Dim wName As String
        Dim wImg As String
        Dim wImghw As String
        Dim wX As String
        Dim wY As String
        Dim wXOffset As String
        Dim wYOffset As String



        '===========種類=============
        If RadioButton1.Checked = True Then
            wType = 0   'クロスヘア
        ElseIf RadioButton2.Checked = True Then
            wType = 1   'ドット
        ElseIf RadioButton3.Checked = True Then
            wType = 2   'ドット＆サークル
        ElseIf RadioButton6.Checked = True Then
            wType = 3   '画像
        End If

        '===========大きさ=============
        If RadioButton4.Checked = True Then
            wSize = 0   'Medium
        ElseIf RadioButton5.Checked = True Then
            wSize = 1   'Large
        End If

        '===========色=============
        wColor = TextBox1.Text & "," & TextBox2.Text & "," & TextBox3.Text


        '===========名前=============
        wName = TextBox4.Text

        '===========画像=============
        wImg = TextBox5.Text

        '===========画像の設定=============
        If RadioButton7.Checked = True Then
            '縮小
            wImghw = "0,0,0"
        ElseIf RadioButton8.Checked Then
            '原寸大
            wImghw = "1,0,0"
        Else
            '指定
            wImghw = "3," & NumericUpDown3.Text & "," & NumericUpDown4.Text
        End If

        '===========x軸y軸=============
        wY = NumericUpDown1.Value
        wX = NumericUpDown2.Value
        wYOffset = NumericUpDown5.Value
        wXOffset = NumericUpDown6.Value


        '画像ファイルを読み込んで、Imageオブジェクトとして取得する
        If wType = 3 Then

            Try
                Image.FromFile(wImg)
            Catch ex As System.OutOfMemoryException
                MsgBox("ファイルが正しくありません。")
                Return
            Catch ex As System.ArgumentException
                MsgBox("ファイルの形式が無効です。")
                Return
            Catch ex As System.IO.FileNotFoundException
                MsgBox("ファイルが存在しません。")
                Return
            End Try
        End If

        'それぞれ対応したセッティングに入れる
        If ListBox1.SelectedIndex = 0 Then
            My.Settings.type1 = wType
            My.Settings.size1 = wSize
            My.Settings.color1 = wColor
            My.Settings.name1 = wName
            My.Settings.img1 = wImg
            My.Settings.imghw1 = wImghw
            My.Settings.x1 = wX
            My.Settings.y1 = wY
            My.Settings.x_offset1 = wXOffset
            My.Settings.y_offset1 = wYOffset
        ElseIf ListBox1.SelectedIndex = 1 Then
            My.Settings.type2 = wType
            My.Settings.size2 = wSize
            My.Settings.color2 = wColor
            My.Settings.name2 = wName
            My.Settings.img2 = wImg
            My.Settings.imghw2 = wImghw
            My.Settings.x2 = wX
            My.Settings.y2 = wY
            My.Settings.x_offset2 = wXOffset
            My.Settings.y_offset2 = wYOffset
        ElseIf ListBox1.SelectedIndex = 2 Then
            My.Settings.type3 = wType
            My.Settings.size3 = wSize
            My.Settings.color3 = wColor
            My.Settings.name3 = wName
            My.Settings.img3 = wImg
            My.Settings.imghw3 = wImghw
            My.Settings.x3 = wX
            My.Settings.y3 = wY
            My.Settings.x_offset3 = wXOffset
            My.Settings.y_offset3 = wYOffset
        ElseIf ListBox1.SelectedIndex = 3 Then
            My.Settings.type4 = wType
            My.Settings.size4 = wSize
            My.Settings.color4 = wColor
            My.Settings.name4 = wName
            My.Settings.img4 = wImg
            My.Settings.imghw4 = wImghw
            My.Settings.x4 = wX
            My.Settings.y4 = wY
            My.Settings.x_offset4 = wXOffset
            My.Settings.y_offset4 = wYOffset
        ElseIf ListBox1.SelectedIndex = 4 Then
            My.Settings.type5 = wType
            My.Settings.size5 = wSize
            My.Settings.color5 = wColor
            My.Settings.name5 = wName
            My.Settings.img5 = wImg
            My.Settings.imghw5 = wImghw
            My.Settings.x5 = wX
            My.Settings.y5 = wY
            My.Settings.x_offset5 = wXOffset
            My.Settings.y_offset5 = wYOffset
        End If



        'リストボックスの更新
        ldsetName()

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        Dim loadType As Integer
        Dim loadSize As Integer
        Dim loadColor As String
        Dim loadImg As String
        Dim loadImghw As String
        Dim loadX As String
        Dim loadY As String
        Dim loadXOffset As String
        Dim loadYOffset As String

        If Not ListBox1.SelectedIndex = -1 Then
            If ListBox1.SelectedIndex = 0 Then
                loadType = My.Settings.type1
                loadSize = My.Settings.size1
                loadColor = My.Settings.color1
                loadImg = My.Settings.img1
                loadImghw = My.Settings.imghw1
                loadX = My.Settings.x1
                loadY = My.Settings.y1
                loadXOffset = My.Settings.x_offset1
                loadYOffset = My.Settings.y_offset1
            ElseIf ListBox1.SelectedIndex = 1 Then
                loadType = My.Settings.type2
                loadSize = My.Settings.size2
                loadColor = My.Settings.color2
                loadImg = My.Settings.img2
                loadImghw = My.Settings.imghw2
                loadX = My.Settings.x2
                loadY = My.Settings.y2
                loadXOffset = My.Settings.x_offset2
                loadYOffset = My.Settings.y_offset2
            ElseIf ListBox1.SelectedIndex = 2 Then
                loadType = My.Settings.type3
                loadSize = My.Settings.size3
                loadColor = My.Settings.color3
                loadImg = My.Settings.img3
                loadImghw = My.Settings.imghw3
                loadX = My.Settings.x3
                loadY = My.Settings.y3
                loadXOffset = My.Settings.x_offset3
                loadYOffset = My.Settings.y_offset3
            ElseIf ListBox1.SelectedIndex = 3 Then
                loadType = My.Settings.type4
                loadSize = My.Settings.size4
                loadColor = My.Settings.color4
                loadImg = My.Settings.img4
                loadImghw = My.Settings.imghw4
                loadX = My.Settings.x4
                loadY = My.Settings.y4
                loadXOffset = My.Settings.x_offset4
                loadYOffset = My.Settings.y_offset4
            ElseIf ListBox1.SelectedIndex = 4 Then
                loadType = My.Settings.type5
                loadSize = My.Settings.size5
                loadColor = My.Settings.color5
                loadImg = My.Settings.img5
                loadImghw = My.Settings.imghw5
                loadX = My.Settings.x5
                loadY = My.Settings.y5
                loadXOffset = My.Settings.x_offset5
                loadYOffset = My.Settings.y_offset5
            End If



            '===========種類=============
            If loadType = 0 Then
                RadioButton1.Checked = True
            ElseIf loadType = 1 Then
                RadioButton2.Checked = True
            ElseIf loadType = 2 Then
                RadioButton3.Checked = True
            ElseIf loadType = 3 Then
                RadioButton6.Checked = True
            End If

            '===========大きさ=============
            If loadSize = 0 Then
                RadioButton4.Checked = True
            ElseIf loadSize = 1 Then
                RadioButton5.Checked = True
            End If

            '===========色=============
            ' カンマ区切りで分割して配列に格納する
            Dim stArrayData As String() = Split(loadColor, ",")
            TextBox1.Text = stArrayData(0)
            TextBox2.Text = stArrayData(1)
            TextBox3.Text = stArrayData(2)


            '===========画像の設定=============
            Dim imgArrayData As String() = Split(loadImghw, ",")

            '画像
            If Not loadImg = "" Then
                ofd.FileName = loadImg
                showpre()

                Try
                    If imgArrayData(0) = 3 Then 'NumericUpDownに画像の大きさを設定する
                        NumericUpDown3.Value = imgArrayData(1)
                        NumericUpDown4.Value = imgArrayData(2)
                    Else

                        NumericUpDown3.Value = oriw
                        NumericUpDown4.Value = orih
                    End If

                Catch ex As System.ArgumentOutOfRangeException

                    MsgBox("ファイルサイズが正しくありません。", MsgBoxStyle.Exclamation)

                End Try
            End If

            If imgArrayData(0) = 0 Then
                RadioButton7.Checked = True
            ElseIf imgArrayData(0) = 1 Then
                RadioButton8.Checked = True
            Else
                RadioButton9.Checked = True
            End If

            '===========x軸y軸=============
            NumericUpDown1.Value = loadY
            NumericUpDown2.Value = loadX
            NumericUpDown5.Value = loadYOffset
            NumericUpDown6.Value = loadXOffset


        End If
    End Sub

    Private Function settingImg(path, type)

        If path = "" Then
            PictureBox3.Image = Nothing
            Label20.Text = ""
            Return Nothing
        Else
            Label20.Text = ""
        End If

        Dim setFileName As String = path
        Dim stw As Integer = 0
        Dim sth As Integer = 0

        Try
            img = Image.FromFile(path)
        Catch ex As System.OutOfMemoryException
            MsgBox("ファイルが正しくありません。")
            Return Nothing
        Catch ex As System.ArgumentException
            MsgBox("ファイルの形式が無効です。")
            Return Nothing
        Catch ex As System.IO.FileNotFoundException
            MsgBox("ファイルが存在しません。")
            Return Nothing
        End Try

        oriw = img.Width
        orih = img.Height

        'PictureBox3の大きさに合わせる計算
        If (60 < img.Width Or 60 < img.Height) Then
            PictureBox3.Width = 60
            PictureBox3.Height = 60
            If (img.Width >= img.Height) Then
                stw = img.Width / (img.Width / 60)
                sth = img.Height / (img.Width / 60)
            Else
                stw = img.Width / (img.Height / 60)
                sth = img.Height / (img.Height / 60)

            End If
        Else
            stw = img.Width
            sth = img.Height
        End If

        Dim canvas As New Bitmap(stw, sth)
        'ImageオブジェクトのGraphicsオブジェクトを作成する
        Dim stg As Graphics = Graphics.FromImage(canvas)

        '画像をcanvasの座標(0, 0)の位置に描画する
        stg.DrawImage(img, 0, 0, stw, sth)
        'PictureBox3に表示する
        PictureBox3.Image = canvas

        'Imageオブジェクトのリソースを解放する
        img.Dispose()

        'Graphicsオブジェクトのリソースを解放する
        stg.Dispose()

    End Function

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

        'リスト外をクリックしたときは何もしない
        If ListBox1.SelectedIndex = -1 Then
            Return
        End If

        Dim pType As String
        Dim pSize As String
        Dim pColor As String
        Dim pImg As String
        Dim pX As String
        Dim pY As String


        If ListBox1.SelectedIndex = 0 Then
            pType = My.Settings.type1
            pSize = My.Settings.size1
            pColor = My.Settings.color1
            TextBox4.Text = My.Settings.name1
            pImg = My.Settings.img1
            pX = My.Settings.x1
            pY = My.Settings.y1
        ElseIf ListBox1.SelectedIndex = 1 Then
            pType = My.Settings.type2
            pSize = My.Settings.size2
            pColor = My.Settings.color2
            TextBox4.Text = My.Settings.name2
            pImg = My.Settings.img2
            pX = My.Settings.x2
            pY = My.Settings.y2
        ElseIf ListBox1.SelectedIndex = 2 Then
            pType = My.Settings.type3
            pSize = My.Settings.size3
            pColor = My.Settings.color3
            TextBox4.Text = My.Settings.name3
            pImg = My.Settings.img3
            pX = My.Settings.x3
            pY = My.Settings.y3
        ElseIf ListBox1.SelectedIndex = 3 Then
            pType = My.Settings.type4
            pSize = My.Settings.size4
            pColor = My.Settings.color4
            TextBox4.Text = My.Settings.name4
            pImg = My.Settings.img4
            pX = My.Settings.x4
            pY = My.Settings.y4
        ElseIf ListBox1.SelectedIndex = 4 Then
            pType = My.Settings.type5
            pSize = My.Settings.size5
            pColor = My.Settings.color5
            TextBox4.Text = My.Settings.name5
            pImg = My.Settings.img5
            pX = My.Settings.x5
            pY = My.Settings.y5
        End If


        '===========種類=============
        If pType = 0 Then
            Label14.Text = "クロスヘア"
        ElseIf pType = 1 Then
            Label14.Text = "ドット"
        ElseIf pType = 2 Then
            Label14.Text = "ドット＆サークル"
        ElseIf pType = 3 Then
            Label14.Text = "画像"
        End If

        '===========大きさ=============
        If pSize = 0 Then
            Label16.Text = "Medium"
        ElseIf pSize = 1 Then
            Label16.Text = "Large"
        End If

        '===========色=============
        Dim stArrayData As String() = Split(pColor, ",")
        Panel3.BackColor = Color.FromArgb(stArrayData(0), stArrayData(1), stArrayData(2))

        '===========画像=============
        settingImg(pImg, pType)

        '===========x軸y軸=============
        NumericUpDown7.Value = pY
        NumericUpDown8.Value = pX


    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        ofd.FileName = TextBox5.Text
        showpre()

        Try
            'NumericUpDownに画像の大きさを設定する
            NumericUpDown3.Value = oriw
            NumericUpDown4.Value = orih

        Catch ex As System.ArgumentOutOfRangeException

            MsgBox("ファイルサイズが正しくありません。", MsgBoxStyle.Exclamation)

        End Try

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        My.Settings.def_set = ListBox1.SelectedIndex
        MsgBox("起動時のデフォルトを「" & TextBox4.Text & "」に設定しました")
    End Sub


    Private Sub Form1_FormClosing(ByVal sender As System.Object,
     ByVal e As System.Windows.Forms.FormClosingEventArgs) _
     Handles MyBase.FormClosing
        If MouseHook.Hooked = True Then
            MouseHook.MouseHookEnd()
        End If

    End Sub

    Private Sub setWin()
        SetWindowLong(Form2.Handle, GWL_EXSTYLE, GetWindowLong(Form2.Handle, GWL_EXSTYLE) Xor WS_EX_LAYERED Xor WS_EX_TRANSPARENT)

        SetLayeredWindowAttributes(Form2.Handle, RGB(255, 0, 0), 255, LWA_COLORKEY)
        SetWindowPos(Form2.Handle, HWND_TOPMOST, 0, 0, 0, 0,
            TOPMOST_FLAGS)
    End Sub



    Private Sub MouseHook_MouseHook(sender As Object, e As MouseHookClass.MouseHookEventArgs) Handles MouseHook.MouseHook
        Dim mStat As String

        mStat = String.Format("{0}", e.Message)
            If stat = 1 Then
                If mStat = "RDown" Then
                    Form2.Hide()
                ElseIf mStat = "RUp" Then
                    Form2.Show()
                End If
            End If


    End Sub

    Private Sub CheckBox3_Checked(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged

        Select Case DirectCast(sender, CheckBox).CheckState
            Case CheckState.Checked
                CheckBox2.Enabled = False
            Case CheckState.Unchecked
                CheckBox2.Enabled = True
        End Select

        If MouseHook.Hooked = False Then
                If MouseHook.MouseHookStart() = True Then

                End If
            Else
                If MouseHook.MouseHookEnd() = True Then

                End If
            End If

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        'チェック状態の確認
        Select Case DirectCast(sender, CheckBox).CheckState
            Case CheckState.Checked
                CheckBox3.Enabled = False
            Case CheckState.Unchecked
                CheckBox3.Enabled = True
        End Select
    End Sub


End Class


Public Delegate Function CallBack(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer

Public Class MouseHookClass

    Dim WH_MOUSE_LL As Integer = 14
    Shared hHook As Integer = 0

    Private hookproc As CallBack

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Overloads Shared Function GetModuleHandle(lpModuleName As String) As IntPtr
    End Function

    'Import for the SetWindowsHookEx function.
    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Overloads Shared Function SetWindowsHookEx(ByVal idHook As Integer, ByVal HookProc As CallBack, ByVal hInstance As IntPtr, ByVal wParam As Integer) As Integer
    End Function

    'Import for the CallNextHookEx function.
    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Overloads Shared Function CallNextHookEx(ByVal idHook As Integer, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    End Function
    'Import for the UnhookWindowsHookEx function.
    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Overloads Shared Function UnhookWindowsHookEx(ByVal idHook As Integer) As Boolean
    End Function

    'Point structure declaration.
    <StructLayout(LayoutKind.Sequential)> Public Structure Point
        Public x As Integer
        Public y As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Class MouseLLHookStruct
        Public pt As Point
        Public mouseData As Integer
        Public flags As Integer
        Public time As Integer
        Public dwExtraInfo As Integer
    End Class

    'マウス操作の種類を表す。
    Public Enum MouseMessage
        '右ボタンが押された。
        RDown = &H204
        '右ボタンが解放された。
        RUp = &H205

    End Enum


    Public Event MouseHook(sender As Object, e As MouseHookEventArgs)
    Public Class MouseHookEventArgs
        Inherits EventArgs

        Private _mousestatus As MouseLLHookStruct
        Private _mousemessage As MouseMessage
        Public Sub New(mousemessage As MouseMessage, mousestatus As MouseLLHookStruct)
            _mousemessage = mousemessage
            _mousestatus = mousestatus
        End Sub


        ''' <summary>
        ''' マウスの状態
        ''' </summary>
        Public ReadOnly Property Message As MouseMessage
            Get
                Return _mousemessage
            End Get
        End Property
    End Class


    ''' <summary>
    ''' 現在マウスをフックしているか返す
    ''' </summary>
    ''' <returns>False:フックしていない  True:フックしている</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Hooked As Boolean
        Get
            Return If(hHook = 0, False, True)
        End Get
    End Property

    ''' <summary>
    ''' マウスフックを開始する
    ''' </summary>
    ''' <returns>False:フックに失敗もしくはフック済み True:フックに成功</returns>
    ''' <remarks></remarks>
    Public Function MouseHookStart() As Boolean
        If hHook.Equals(0) Then
            'マウスフックを開始する
            hookproc = AddressOf MouseLLHookProc
            hHook = SetWindowsHookEx(WH_MOUSE_LL, hookproc, GetModuleHandle(IntPtr.Zero), 0)
            If hHook.Equals(0) Then
                Return False
            Else
                Return True
            End If
        Else
            'マウスフックがすでに開始されている
            Return False
        End If

    End Function

    ''' <summary>
    ''' マウスフックを終了する
    ''' </summary>
    ''' <returns>False:フック解除に失敗もしくはフックしていない True:フック解除に成功</returns>
    ''' <remarks></remarks>
    Public Function MouseHookEnd() As Boolean
        If hHook.Equals(0) Then
            'マウスフックが開始されていない
            Return False
        Else
            'マウスフックを終了する
            Dim ret As Boolean = UnhookWindowsHookEx(hHook)

            If ret.Equals(False) Then
                Return False
            Else
                hHook = 0
                Return True
            End If
        End If

    End Function

    Private Function MouseLLHookProc(ByVal nCode As Integer, ByVal wParam As MouseMessage, ByVal lParam As IntPtr) As Integer
        Dim MyMouseHookStruct As New MouseLLHookStruct()

        If nCode = 0 Then
            MyMouseHookStruct = CType(Marshal.PtrToStructure(lParam, MyMouseHookStruct.GetType()), MouseLLHookStruct)
            'イベントを発生させる
            RaiseEvent MouseHook(Nothing, New MouseHookEventArgs(wParam, MyMouseHookStruct))
        End If

        Return CallNextHookEx(hHook, nCode, wParam, lParam)
    End Function

End Class