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

    Public Declare Function FindWindowA Lib "user32" (ByVal cnm As String, ByVal cap As String) As IntPtr

    Dim stat As Integer
    Dim count As Integer

    Dim Rect2 As Rectangle
    Dim DisWidth As Integer
    Dim DisHeight As Integer
    Dim FirstWidth As Integer
    Dim FirstHeight As Integer
    Dim NumWidth As Integer
    Dim NumHeight As Integer

    Dim colorR As Integer
    Dim colorG As Integer
    Dim colorB As Integer

    '1 = 画像
    Dim statImage As Integer

    Dim fWait As Integer

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

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'カーソル位置のスクリーンを基準にする（マルチモニタ対応）
        Rect2 = Screen.FromPoint(Cursor.Position).Bounds
        'ポジションに設定
        DisHeight = Rect2.Top + Rect2.Height / 2 - (Form2.Height / 2)
        DisWidth = Rect2.Left + Rect2.Width / 2 - (Form2.Width / 2)

        statImage = 0

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

        ListBox1.SelectedIndex = My.Settings.def_set
        Button11_Click(Nothing, Nothing)

        SetWindowLong(Form2.Handle, GWL_EXSTYLE, GetWindowLong(Form2.Handle, GWL_EXSTYLE) Or WS_EX_LAYERED Or WS_EX_TRANSPARENT)

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
        For i As Integer = 1 To PresetCount
            ListBox1.Items.Add(CStr(My.Settings.Item("name" & i)))
        Next
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

    '現在の種類選択を形状名で返す。"cross"/"dot"/"circle"/"image"
    Private Function CurrentShape() As String
        If RadioButton1.Checked Then
            Return "cross"
        ElseIf RadioButton2.Checked Then
            Return "dot"
        ElseIf RadioButton3.Checked Then
            Return "circle"
        End If
        Return "image"
    End Function

    '現在の大きさ選択をサイズ名で返す。"small"/"large"
    Private Function CurrentSize() As String
        If RadioButton4.Checked Then
            Return "small"
        End If
        Return "large"
    End Function

    '現在の選択でオーバーレイを作り直して表示する。
    Private Async Function RefreshOverlay() As Task
        If CurrentShape() = "image" Then
            showimage()
        Else
            Await crosshair(CurrentSize(), CurrentShape())
        End If
        Form2.Show()
        Form2.Location = New Point(DisWidth, DisHeight)
    End Function

    '現在の選択でプレビューを更新する。
    Private Sub UpdatePreview()
        If TabControl1.SelectedIndex <> 0 Then
            f.Hide()
            Return
        End If
        If RadioButton1.Checked Then
            If RadioButton4.Checked Then preview1() Else bigpreview1()
        ElseIf RadioButton2.Checked Then
            If RadioButton4.Checked Then preview2() Else bigpreview2()
        ElseIf RadioButton3.Checked Then
            If RadioButton4.Checked Then preview3() Else bigpreview3()
        End If
    End Sub

    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Button1.Enabled = False
        Button2.Enabled = True
        '起動中ならば1
        stat = 1

        x_axis()
        y_axis()

        Await RefreshOverlay()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Button1.Enabled = True
        Button2.Enabled = False
        '停止中ならば0
        stat = 0
        Form2.Hide()

    End Sub


    Private Async Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If statImage = 1 Then
            setWin()
        End If

        statImage = 0
        If stat = 1 AndAlso RadioButton1.Checked Then
            '起動中ならば表示
            Await RefreshOverlay()
        End If

        UpdatePreview()

    End Sub

    Private Async Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If statImage = 1 Then
            setWin()
        End If
        statImage = 0
        If stat = 1 AndAlso RadioButton2.Checked Then
            Await RefreshOverlay()
        End If

        UpdatePreview()

    End Sub

    Private Async Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If statImage = 1 Then
            setWin()
        End If

        statImage = 0

        If stat = 1 AndAlso RadioButton3.Checked Then
            Await RefreshOverlay()
        End If


        UpdatePreview()



    End Sub

    Private Async Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged
        If RadioButton6.Checked Then
            '画像の描写
            statImage = 1
            Form2.PictureBox1.Visible = True

            If stat = 1 Then
                Await RefreshOverlay()
            End If
        Else
            Form2.PictureBox1.Visible = False
        End If

    End Sub

    '形状名("cross"/"dot"/"circle")とサイズ("small"/"large")からGraphicsPathを生成する。
    'プレビューと実表示で座標を共有するための単一ソース。
    Private Function BuildShapePath(shape As String, size As String) As GraphicsPath
        Dim path As New GraphicsPath()
        path.StartFigure()

        If shape = "cross" Then 'クロスヘア
            If size = "small" Then
                path.AddLines(
                {New Point(17, 16),
                New Point(2, 17),
                New Point(17, 18),
                New Point(17, 32),
                New Point(18, 18),
                New Point(32, 17),
                New Point(18, 16),
                New Point(17, 2)})
            Else
                path.AddLines(
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
                New Point(16, 2)})
            End If
        ElseIf shape = "dot" Then
            If size = "small" Then
                path.AddEllipse(New Rectangle(15, 15, 4, 4))
            Else
                path.AddEllipse(New Rectangle(13, 13, 8, 8))
            End If
        ElseIf shape = "circle" Then
            If size = "small" Then
                path.AddEllipse(New Rectangle(2, 2, 30, 30))
                path.AddEllipse(New Rectangle(3, 3, 28, 28))
                path.AddEllipse(New Rectangle(15, 15, 4, 4))
            Else
                path.AddEllipse(New Rectangle(0, 0, 34, 34))
                path.AddEllipse(New Rectangle(2, 2, 30, 30))
                path.AddEllipse(New Rectangle(13, 13, 8, 8))
            End If
        End If

        Return path
    End Function

    'フォームのRegionを差し替え、旧Regionを破棄する。
    Private Sub ApplyRegion(target As Form, path As GraphicsPath)
        Dim oldRegion As Region = target.Region
        target.Region = New Region(path)
        If oldRegion IsNot Nothing Then
            oldRegion.Dispose()
        End If
    End Sub

    'プレビュー用フォームfに形状を反映する。
    Private Sub ShowPreview(path As GraphicsPath)
        ApplyRegion(f, path)
        'TopLevelをFalseにする
        f.TopLevel = False
        'フォームのコントロールに追加する
        If Not Me.Controls.Contains(f) Then
            Me.Controls.Add(f)
        End If
        'フォームを表示する
        If TabControl1.SelectedIndex = 0 Then
            f.Show()
        End If
        f.Location = New Point(150, 216)
        '最前面へ移動
        f.BringToFront()
    End Sub

    Private Async Function crosshair(ByVal sizes_in As String, ByVal types_in As String) As Task

        If CheckBox2.Checked Then
            Dim parsedWait As Integer
            If Integer.TryParse(TextBox6.Text, parsedWait) Then
                If parsedWait < 0 Then parsedWait = 0
                If parsedWait > 60 Then parsedWait = 60
                fWait = parsedWait
                Await System.Threading.Tasks.Task.Delay(fWait * 1000)
            End If
        End If

        Using path As GraphicsPath = BuildShapePath(types_in, sizes_in)
            ApplyRegion(Form2, path)
        End Using
    End Function

    Private Sub showimage()

        setWin()


        Dim points() As Point = {}


        'GraphicsPathの作成
        Using path As New GraphicsPath
            path.StartFigure()

            points =
            {New Point(0, 0),
            New Point(0, h),
            New Point(w, h),
            New Point(w, 0)}

            path.AddLines(points)


            ApplyRegion(Form2, path)
        End Using
        Form2.PictureBox1.Image = PictureBox2.Image


        If w > 0 And h > 0 Then
            Form2.Size = New Size(w, h)
            Form2.PictureBox1.Width = w
            Form2.PictureBox1.Height = h
        End If



    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        UpdatePreview()
    End Sub

    Private Sub preview1()
        '---------クロスヘアプレビュー用---------
        Dim path As GraphicsPath = BuildShapePath("cross", "small")

        ShowPreview(path)
        path.Dispose()
    End Sub

    Private Sub preview2()
        Dim path As GraphicsPath = BuildShapePath("dot", "small")

        ShowPreview(path)
        path.Dispose()
    End Sub

    Private Sub preview3()
        Dim path As GraphicsPath = BuildShapePath("circle", "small")

        ShowPreview(path)
        path.Dispose()
    End Sub

    Private Sub bigpreview1()
        Dim path As GraphicsPath = BuildShapePath("cross", "large")

        ShowPreview(path)
        path.Dispose()
    End Sub

    Private Sub bigpreview2()

        Dim path As GraphicsPath = BuildShapePath("dot", "large")

        ShowPreview(path)
        path.Dispose()
    End Sub

    Private Sub bigpreview3()

        Dim path As GraphicsPath = BuildShapePath("circle", "large")

        ShowPreview(path)
        path.Dispose()
    End Sub

    Private Async Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        'Medium

        UpdatePreview()
        If stat = 1 AndAlso statImage = 0 AndAlso RadioButton4.Checked Then
            Await RefreshOverlay()
        End If
    End Sub

    Private Async Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        'large

        UpdatePreview()
        If stat = 1 AndAlso statImage = 0 AndAlso RadioButton5.Checked Then
            Await RefreshOverlay()
        End If
    End Sub



    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Dim v1 As Integer
        If Not Integer.TryParse(TextBox1.Text, v1) Then
            Return
        End If
        If v1 > 255 Then
            v1 = 255
            TextBox1.Text = "255"
        ElseIf v1 < 0 Then
            v1 = 0
        End If
        colorR = v1
        colorset()
    End Sub


    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Dim v2 As Integer
        If Not Integer.TryParse(TextBox2.Text, v2) Then
            Return
        End If
        If v2 > 255 Then
            v2 = 255
            TextBox2.Text = "255"
        ElseIf v2 < 0 Then
            v2 = 0
        End If
        colorG = v2
        colorset()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Dim v3 As Integer
        If Not Integer.TryParse(TextBox3.Text, v3) Then
            Return
        End If
        If v3 > 255 Then
            v3 = 255
            TextBox3.Text = "255"
        ElseIf v3 < 0 Then
            v3 = 0
        End If
        colorB = v3
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
        NumHeight = FirstHeight - CInt(NumericUpDown1.Value + NumericUpDown5.Value)
        DisHeight = NumHeight
        Form2.Location = New Point(DisWidth, NumHeight)

    End Sub

    Public Sub x_axis()
        NumWidth = FirstWidth + CInt(NumericUpDown2.Value + NumericUpDown6.Value)
        DisWidth = NumWidth
        Form2.Location = New Point(NumWidth, DisHeight)
    End Sub

    '設定値の文字列をNumericUpDownに安全に反映する。不正値・範囲外はクランプする。
    Private Sub SetAxisValue(nud As NumericUpDown, text As String)
        Dim v As Decimal
        If Decimal.TryParse(CStr(text), v) Then
            nud.Value = Math.Max(nud.Minimum, Math.Min(nud.Maximum, v))
        Else
            nud.Value = 0
        End If
    End Sub

    '数値をNumericUpDownに範囲内で反映する。
    Private Sub SetAxisValue(nud As NumericUpDown, v As Decimal)
        nud.Value = Math.Max(nud.Minimum, Math.Min(nud.Maximum, v))
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
        Dim hwindow As IntPtr
        Dim x = ComboBox1.Text
        hwindow = FindWindowA(vbNullString, x)
        Dim nRect As RECT

        If hwindow = IntPtr.Zero Then
            If Not x = "" Then
                MsgBox("取得失敗")
            End If
            Return
        End If

        Call GetWindowRect(hwindow, nRect)
        If nRect.Top = -32000 Then
            MsgBox("最小化されています")
        ElseIf nRect.Top = 0 And nRect.Bottom = 0 And nRect.Right = 0 And nRect.Left = 0 Then
            MsgBox("取得失敗")
        ElseIf x = "" Then
        Else
            DisHeight = (nRect.Bottom + SystemInformation.CaptionHeight - nRect.Top) / 2 + nRect.Top - 17
            DisWidth = (nRect.Right - nRect.Left) / 2 + nRect.Left - 17
            SetAxisValue(NumericUpDown1, FirstHeight - DisHeight)
            SetAxisValue(NumericUpDown2, DisWidth - FirstWidth)
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        'リンク先に移動したことにする
        LinkLabel1.LinkVisited = True
        'ブラウザで開く
        System.Diagnostics.Process.Start("https://mjh.blog.jp/")
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
    Private Sub showpre()

        '画像ファイルを読み込んで、Imageオブジェクトとして取得する
        Try
            img = Image.FromFile(ofd.FileName)
        Catch ex As System.OutOfMemoryException
            MsgBox("ファイルが正しくありません。", MsgBoxStyle.Exclamation)
            Return
        Catch ex As System.ArgumentException
            MsgBox("ファイルの形式が無効です。", MsgBoxStyle.Exclamation)
            Return
        Catch ex As System.IO.FileNotFoundException
            MsgBox("ファイルが存在しません。", MsgBoxStyle.Exclamation)
            Return
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

        If w <= 0 OrElse h <= 0 Then
            img.Dispose()
            Return
        End If

        Dim oldPreview As Image = PictureBox2.Image
        Using canvas As New Bitmap(w, h)
            'ImageオブジェクトのGraphicsオブジェクトを作成する
            Using g As Graphics = Graphics.FromImage(canvas)

                '画像をcanvasの座標(0, 0)の位置に描画する
                g.DrawImage(img, 0, 0, w, h)
            End Using
            'PictureBox2に表示する
            PictureBox2.Image = CType(canvas.Clone(), Image)
        End Using

        'Imageオブジェクトのリソースを解放する
        img.Dispose()

        If oldPreview IsNot Nothing Then
            oldPreview.Dispose()
        End If


        If stat = 1 And RadioButton6.Checked = True Then

            Form2.Hide()
            showimage()
            Form2.Show()
            Form2.Location = New Point(DisWidth, DisHeight)
            'If RadioButton6.Checked = False Then
            '    Form2.Hide()
            'End If
        End If
    End Sub

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

    'プリセット1件分の設定値。My.SettingsのtypeN/sizeN/...群と1対1に対応する。
    Private Class PresetData
        Public PType As Integer
        Public PSize As Integer
        Public PColor As String
        Public PName As String
        Public PImg As String
        Public PImghw As String
        Public PX As Integer
        Public PY As Integer
        Public PXOffset As Integer
        Public PYOffset As Integer
    End Class

    Private Const PresetCount As Integer = 5

    Private Function ReadPreset(index As Integer) As PresetData
        Dim d As New PresetData()
        d.PType = CInt(My.Settings.Item("type" & index))
        d.PSize = CInt(My.Settings.Item("size" & index))
        d.PColor = CStr(My.Settings.Item("color" & index))
        d.PName = CStr(My.Settings.Item("name" & index))
        d.PImg = CStr(My.Settings.Item("img" & index))
        d.PImghw = CStr(My.Settings.Item("imghw" & index))
        d.PX = CInt(My.Settings.Item("x" & index))
        d.PY = CInt(My.Settings.Item("y" & index))
        d.PXOffset = CInt(My.Settings.Item("x_offset" & index))
        d.PYOffset = CInt(My.Settings.Item("y_offset" & index))
        Return d
    End Function

    Private Sub WritePreset(index As Integer, d As PresetData)
        My.Settings.Item("type" & index) = d.PType
        My.Settings.Item("size" & index) = d.PSize
        My.Settings.Item("color" & index) = d.PColor
        My.Settings.Item("name" & index) = d.PName
        My.Settings.Item("img" & index) = d.PImg
        My.Settings.Item("imghw" & index) = d.PImghw
        My.Settings.Item("x" & index) = d.PX
        My.Settings.Item("y" & index) = d.PY
        My.Settings.Item("x_offset" & index) = d.PXOffset
        My.Settings.Item("y_offset" & index) = d.PYOffset
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        Dim d As New PresetData()



        '===========種類=============
        If RadioButton1.Checked = True Then
            d.PType = 0   'クロスヘア
        ElseIf RadioButton2.Checked = True Then
            d.PType = 1   'ドット
        ElseIf RadioButton3.Checked = True Then
            d.PType = 2   'ドット＆サークル
        ElseIf RadioButton6.Checked = True Then
            d.PType = 3   '画像
        End If

        '===========大きさ=============
        If RadioButton4.Checked = True Then
            d.PSize = 0   'Medium
        ElseIf RadioButton5.Checked = True Then
            d.PSize = 1   'Large
        End If

        '===========色=============
        d.PColor = TextBox1.Text & "," & TextBox2.Text & "," & TextBox3.Text


        '===========名前=============
        d.PName = TextBox4.Text

        '===========画像=============
        d.PImg = TextBox5.Text

        '===========画像の設定=============
        If RadioButton7.Checked = True Then
            '縮小
            d.PImghw = "0,0,0"
        ElseIf RadioButton8.Checked Then
            '原寸大
            d.PImghw = "1,0,0"
        Else
            '指定
            d.PImghw = "3," & NumericUpDown3.Text & "," & NumericUpDown4.Text
        End If

        '===========x軸y軸=============
        d.PY = CInt(NumericUpDown1.Value)
        d.PX = CInt(NumericUpDown2.Value)
        d.PYOffset = CInt(NumericUpDown5.Value)
        d.PXOffset = CInt(NumericUpDown6.Value)


        '画像ファイルを読み込んで、Imageオブジェクトとして取得する
        If d.PType = 3 Then

            Try
                Using tmp As Image = Image.FromFile(d.PImg)
                End Using
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
        If ListBox1.SelectedIndex >= 0 AndAlso ListBox1.SelectedIndex <= 4 Then
            WritePreset(ListBox1.SelectedIndex + 1, d)
        End If



        'リストボックスの更新
        ldsetName()

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        If ListBox1.SelectedIndex < 0 OrElse ListBox1.SelectedIndex > PresetCount - 1 Then
            Return
        End If

        Dim d As PresetData = ReadPreset(ListBox1.SelectedIndex + 1)
        Dim loadType As Integer = d.PType
        Dim loadSize As Integer = d.PSize
        Dim loadColor As String = d.PColor
        Dim loadImg As String = d.PImg
        Dim loadImghw As String = d.PImghw
        Dim loadX As String = CStr(d.PX)
        Dim loadY As String = CStr(d.PY)
        Dim loadXOffset As String = CStr(d.PXOffset)
        Dim loadYOffset As String = CStr(d.PYOffset)



            '===========種類=============
            If loadType = 0 Then
                RadioButton1.Checked = True
            ElseIf loadType = 1 Then
                RadioButton2.Checked = True
            ElseIf loadType = 2 Then
                RadioButton3.Checked = True
            ElseIf loadType = 3 Then
                RadioButton6.Checked = True
            Else
                RadioButton1.Checked = True
            End If

            '===========大きさ=============
            If loadSize = 0 Then
                RadioButton4.Checked = True
            ElseIf loadSize = 1 Then
                RadioButton5.Checked = True
            Else
                RadioButton4.Checked = True
            End If

            '===========色=============
            ' カンマ区切りで分割して配列に格納する
            Dim stArrayData As String() = Split(CStr(loadColor), ",")
            If stArrayData.Length >= 3 Then
                TextBox1.Text = stArrayData(0).Trim()
                TextBox2.Text = stArrayData(1).Trim()
                TextBox3.Text = stArrayData(2).Trim()
            Else
                TextBox1.Text = "255"
                TextBox2.Text = "0"
                TextBox3.Text = "0"
            End If


            '===========画像の設定=============
            Dim imgArrayData As String() = Split(CStr(loadImghw), ",")
            Dim imgMode As Integer = 0
            If imgArrayData.Length >= 1 Then
                Integer.TryParse(imgArrayData(0).Trim(), imgMode)
            End If

            '画像
            If Not loadImg = "" Then
                ofd.FileName = loadImg
                showpre()

                Try
                    If imgMode = 3 AndAlso imgArrayData.Length >= 3 Then 'NumericUpDownに画像の大きさを設定する
                        Dim lw As Decimal
                        Dim lh As Decimal
                        If Decimal.TryParse(imgArrayData(1).Trim(), lw) Then
                            NumericUpDown3.Value = Math.Max(NumericUpDown3.Minimum, Math.Min(NumericUpDown3.Maximum, lw))
                        End If
                        If Decimal.TryParse(imgArrayData(2).Trim(), lh) Then
                            NumericUpDown4.Value = Math.Max(NumericUpDown4.Minimum, Math.Min(NumericUpDown4.Maximum, lh))
                        End If
                    ElseIf oriw > 0 AndAlso orih > 0 Then

                        NumericUpDown3.Value = Math.Max(NumericUpDown3.Minimum, Math.Min(NumericUpDown3.Maximum, oriw))
                        NumericUpDown4.Value = Math.Max(NumericUpDown4.Minimum, Math.Min(NumericUpDown4.Maximum, orih))
                    End If

                Catch ex As System.ArgumentOutOfRangeException

                    MsgBox("ファイルサイズが正しくありません。", MsgBoxStyle.Exclamation)

                End Try
            End If

            If imgMode = 0 Then
                RadioButton7.Checked = True
            ElseIf imgMode = 1 Then
                RadioButton8.Checked = True
            Else
                RadioButton9.Checked = True
            End If

            '===========x軸y軸=============
            SetAxisValue(NumericUpDown1, loadY)
            SetAxisValue(NumericUpDown2, loadX)
            SetAxisValue(NumericUpDown5, loadYOffset)
            SetAxisValue(NumericUpDown6, loadXOffset)

    End Sub

    Private Sub settingImg(path As String)

        If path = "" Then
            Dim oldEmpty As Image = PictureBox3.Image
            PictureBox3.Image = Nothing
            If oldEmpty IsNot Nothing Then
                oldEmpty.Dispose()
            End If
            Label20.Text = ""
            Return
        Else
            Label20.Text = ""
        End If

        Dim stw As Integer = 0
        Dim sth As Integer = 0

        Try
            img = Image.FromFile(path)
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

        If stw <= 0 OrElse sth <= 0 Then
            img.Dispose()
            Return
        End If

        Dim oldThumb As Image = PictureBox3.Image
        Using canvas As New Bitmap(stw, sth)
            'ImageオブジェクトのGraphicsオブジェクトを作成する
            Using stg As Graphics = Graphics.FromImage(canvas)
                '画像をcanvasの座標(0, 0)の位置に描画する
                stg.DrawImage(img, 0, 0, stw, sth)
            End Using
            'PictureBox3に表示する
            PictureBox3.Image = CType(canvas.Clone(), Image)
        End Using

        'Imageオブジェクトのリソースを解放する
        img.Dispose()

        If oldThumb IsNot Nothing Then
            oldThumb.Dispose()
        End If

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

        'リスト外をクリックしたときは何もしない
        If ListBox1.SelectedIndex < 0 OrElse ListBox1.SelectedIndex > PresetCount - 1 Then
            Return
        End If

        Dim d As PresetData = ReadPreset(ListBox1.SelectedIndex + 1)
        Dim pType As Integer = d.PType
        Dim pSize As Integer = d.PSize
        Dim pColor As String = d.PColor
        Dim pImg As String = d.PImg
        Dim pX As String = CStr(d.PX)
        Dim pY As String = CStr(d.PY)
        TextBox4.Text = d.PName


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
        Dim stArrayData As String() = Split(CStr(pColor), ",")
        Dim pvR As Integer = 255
        Dim pvG As Integer = 0
        Dim pvB As Integer = 0
        If stArrayData.Length >= 3 Then
            Dim t As Integer
            If Integer.TryParse(stArrayData(0).Trim(), t) Then
                pvR = Math.Max(0, Math.Min(255, t))
            End If
            If Integer.TryParse(stArrayData(1).Trim(), t) Then
                pvG = Math.Max(0, Math.Min(255, t))
            End If
            If Integer.TryParse(stArrayData(2).Trim(), t) Then
                pvB = Math.Max(0, Math.Min(255, t))
            End If
        End If
        Panel3.BackColor = Color.FromArgb(pvR, pvG, pvB)

        '===========画像=============
        settingImg(pImg)

        '===========x軸y軸=============
        SetAxisValue(NumericUpDown7, pY)
        SetAxisValue(NumericUpDown8, pX)


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
        SetWindowLong(Form2.Handle, GWL_EXSTYLE, GetWindowLong(Form2.Handle, GWL_EXSTYLE) Or WS_EX_LAYERED Or WS_EX_TRANSPARENT)

        SetLayeredWindowAttributes(Form2.Handle, RGB(255, 0, 0), 255, LWA_COLORKEY)
        SetWindowPos(Form2.Handle, HWND_TOPMOST, 0, 0, 0, 0,
            TOPMOST_FLAGS)
    End Sub



    Private Sub MouseHook_MouseHook(sender As Object, e As MouseHookClass.MouseHookEventArgs) Handles MouseHook.MouseHook
        If stat = 1 Then
            If e.Message = MouseHookClass.MouseMessage.RDown Then
                Form2.Hide()
            ElseIf e.Message = MouseHookClass.MouseMessage.RUp Then
                Form2.Show()
            End If
        End If

    End Sub

    Private Sub CheckBox3_Checked(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged

        If CheckBox3.Checked Then
            CheckBox2.Enabled = False
            If Not MouseHook.Hooked Then
                MouseHook.MouseHookStart()
            End If
        Else
            CheckBox2.Enabled = True
            If MouseHook.Hooked Then
                MouseHook.MouseHookEnd()
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
    Shared hHook As IntPtr = IntPtr.Zero

    Private hookproc As CallBack

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Overloads Shared Function GetModuleHandle(lpModuleName As IntPtr) As IntPtr
    End Function

    'Import for the SetWindowsHookEx function.
    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Overloads Shared Function SetWindowsHookEx(ByVal idHook As Integer, ByVal HookProc As CallBack, ByVal hInstance As IntPtr, ByVal wParam As Integer) As IntPtr
    End Function

    'Import for the CallNextHookEx function.
    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Overloads Shared Function CallNextHookEx(ByVal idHook As IntPtr, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    End Function
    'Import for the UnhookWindowsHookEx function.
    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Overloads Shared Function UnhookWindowsHookEx(ByVal idHook As IntPtr) As Boolean
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
            Return If(hHook.Equals(IntPtr.Zero), False, True)
        End Get
    End Property

    ''' <summary>
    ''' マウスフックを開始する
    ''' </summary>
    ''' <returns>False:フックに失敗もしくはフック済み True:フックに成功</returns>
    ''' <remarks></remarks>
    Public Function MouseHookStart() As Boolean
        If hHook.Equals(IntPtr.Zero) Then
            'マウスフックを開始する
            hookproc = AddressOf MouseLLHookProc
            hHook = SetWindowsHookEx(WH_MOUSE_LL, hookproc, GetModuleHandle(IntPtr.Zero), 0)
            If hHook.Equals(IntPtr.Zero) Then
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
        If hHook.Equals(IntPtr.Zero) Then
            'マウスフックが開始されていない
            Return False
        Else
            'マウスフックを終了する
            Dim ret As Boolean = UnhookWindowsHookEx(hHook)

            If ret.Equals(False) Then
                Return False
            Else
                hHook = IntPtr.Zero
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