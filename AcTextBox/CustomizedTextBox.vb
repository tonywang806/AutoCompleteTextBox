Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
'**
'* 
'* 全角・半角や大文字・小文字などを無視して、候補を選択できる
'* オートコンプリート機能を実現するTextBoxです。
'*
'* Author:    Tony Wang <tonywang806@gmail.com>
'* Created:   12.03.2018
'* 
'**/
Public Class CustomizedTextBox
    Inherits TextBox

    Private _listBox As ListBox
    Private _isAdded As Boolean
    Private _values() As String
    Private _formerValue As String = String.Empty
    '候補内容
    <Description("マッチングロジックをカスタマイズできる場合、候補のCustomeSourceを設定するString()です。") _
    , DefaultValue(False), Category("カスタム")>
    Public Property AutoCompleteValues As String()
        Get
            Return _values
        End Get
        Set(value As String())
            _values = value
        End Set
    End Property
    '選択した候補アイテム
    Public ReadOnly Property SelectedValues As List(Of String)
        Get
            Dim result() As String = Text.Split(New Char() {" "}, StringSplitOptions.RemoveEmptyEntries)
            Return New List(Of String)(result)
        End Get
    End Property

    Public Sub New()
        InitializeComponent()
        ResetListBox()
    End Sub

    'コンポーネント初期化
    Private Sub InitializeComponent()
        '候補表示用リストの初期化
        _listBox = New ListBox()


        'リストアイテムをカスタマイズため
        _listBox.DrawMode = DrawMode.OwnerDrawVariable
        AddHandler _listBox.DrawItem, AddressOf ListBox_DrawItem
        AddHandler _listBox.MouseDoubleClick, AddressOf ListBox_MouseDoubleClick

        'テキストボクッスへKeyDown、KeyUpイベントを追加
        AddHandler Me.KeyDown, AddressOf this_KeyDown
        AddHandler Me.KeyUp, AddressOf this_KeyUp
    End Sub

    Private Sub this_KeyUp(sender As Object, e As KeyEventArgs)
        '入力内容より候補を更新してから、表示・非表示をする
        UpdateListBox()
    End Sub

    Private Sub this_KeyDown(sender As Object, e As KeyEventArgs)
        Select Case e.KeyCode
            Case Keys.Enter, Keys.Tab
                If _listBox.Visible Then
                    '候補が表示された場合、Enterキー、Tabキーを押したら
                    '選択したアイテムをテキストボクッスに表示
                    'そして、候補リストが非表示
                    Text = _listBox.SelectedItem.ToString()
                    ResetListBox()
                    _formerValue = Text
                    Me.Select(Me.Text.Length, 0)
                    e.Handled = True
                End If
            Case Keys.Down
                If (_listBox.Visible) AndAlso (_listBox.SelectedIndex < _listBox.Items.Count - 1) Then
                    '候補が表示された場合、↓キーを押したら、
                    '候補リストのSelectedIndexを1カウントアップ
                    '※(フォーカスがリストにある場合、リストボックス自体のハンドルで処理)
                    _listBox.SelectedIndex += 1
                    e.Handled = True
                End If
            Case Keys.Up
                If (_listBox.Visible) AndAlso (_listBox.SelectedIndex > 0) Then
                    '候補が表示された場合、↑キーを押したら、
                    '候補リストのSelectedIndexを1減らす
                    '※(フォーカスがリストにある場合、リストボックス自体のハンドルで処理)
                    _listBox.SelectedIndex -= 1
                    e.Handled = True
                End If
        End Select
    End Sub

    Private Sub ListBox_MouseDoubleClick(sender As Object, e As MouseEventArgs)
        If _listBox.Visible Then
            '候補リストにて、ダブルクリックされたアイテムを
            '選択値として戻り、テキストボクッスに表示
            'そして、候補リストが非表示
            Text = _listBox.SelectedItem.ToString()
            ResetListBox()
            _formerValue = Text
            Me.Select(Me.Text.Length, 0)
        End If
    End Sub

    Private Sub ResetListBox()
        '候補リストを非表示
        _listBox.Visible = False
    End Sub

    Private Sub UpdateListBox()
        If Text = _formerValue Then
            '入力内容が変わらない場合
            Return
        End If
        _formerValue = Me.Text


        Dim word As String = Me.Text
        If (_values IsNot Nothing) AndAlso (word.Length > 0) Then
            '入力内容より、候補を選ぶ
            Dim matches() As String = Array.FindAll(Of String)(_values, AddressOf IsSmContainsIgnoreCaseWidthaller)

            If matches.Length > 0 Then
                ShowListBox()
                _listBox.BeginUpdate()
                _listBox.Items.Clear()
                '候補をリストへ追加
                Array.ForEach(matches, Function(x) {_listBox.Items.Add(x)})
                _listBox.SelectedIndex = 0
                _listBox.Height = 0
                _listBox.Width = 0
                Focus()
                'アイテムの内容、数より、候補リストの幅と高さを動態的に設定
                Using graphics As Graphics = _listBox.CreateGraphics()
                    For i As Integer = 0 To _listBox.Items.Count - 1
                        If i < 20 Then
                            _listBox.Height += _listBox.GetItemHeight(i) + 2

                            ' it item width Is larger than the current one
                            ' set it to the New max item width
                            ' GetItemRectangle does Not work for me
                            ' we add a little extra space by using '_'
                            Dim itemWidth As Integer = CType(graphics.MeasureString(_listBox.Items(i).ToString + "_", _listBox.Font).Width, Integer)
                            _listBox.Width = If(itemWidth < _listBox.Width, _listBox.Width, itemWidth)
                            _listBox.Width = If(_listBox.Width < Width, Width, _listBox.Width)
                        End If
                    Next
                End Using

                _listBox.EndUpdate()
            Else
                ResetListBox()
            End If
        Else
            ResetListBox()
        End If

    End Sub

    Private Sub ShowListBox()
        If Not _isAdded Then
            Parent.Controls.Add(_listBox)
            _listBox.Left = Left
            _listBox.Top = Top + Height
            _isAdded = True
        End If
        _listBox.Visible = True
        _listBox.BringToFront()
    End Sub

    Protected Overrides Function IsInputKey(ByVal keyData As Keys) As Boolean
        Select Case keyData
            Case Keys.Tab
                If _listBox.Visible Then
                    Return True
                Else
                    Return False
                End If
            Case Else
                Return MyBase.IsInputKey(keyData)
        End Select
    End Function
    'マッチングかどうかを判断する
    Private Function IsSmContainsIgnoreCaseWidthaller(value As String) As Boolean
        If value Is Nothing Then
            Return False
        End If

        If FirstIndexOfMatching(value) >= 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    '先頭からマッチングしたIndexを戻す
    Private Function FirstIndexOfMatching(value As String) As Integer
        '全角と半角・大文字と小文字・ひらがなとカタカナを無視してマッチングロジック
        Dim ci As System.Globalization.CompareInfo = System.Globalization.CultureInfo.CurrentCulture.CompareInfo
        Return ci.IndexOf(value, Me.Text, System.Globalization.CompareOptions.IgnoreKanaType Or
                        System.Globalization.CompareOptions.IgnoreWidth Or
                        System.Globalization.CompareOptions.IgnoreCase)
    End Function
    'すべてマッチングしたIndexを戻す
    Private Function IndexesOfMatching(value As String) As Integer()
        Dim indexList As New List(Of Integer)

        Dim i As Integer = FirstIndexOfMatching(value)

        If i = -1 Then
            Return indexList.ToArray
        End If

        Dim lastPart As String = value.Substring(i + 1)
        indexList.Add(i)

        If lastPart.Length < Text.Length Then
            Return indexList.ToArray
        Else
            Dim subIndex() As Integer = IndexesOfMatching(lastPart)
            For Each x As Integer In subIndex
                indexList.Add(i + 1 + x)
            Next
            Return indexList.ToArray
        End If
    End Function

    'DrawItemイベントハンドラ
    'マッチング分を描画する
    Private Sub ListBox_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs)
        '背景を描画する
        '項目が選択されている時は強調表示される
        e.DrawBackground()

        'ListBoxが空のときにListBoxが選択されるとe.Indexが-1になる
        If e.Index > -1 Then
            '描画する文字列の取得
            Dim currentTxt As String = CType(sender, ListBox).Items(e.Index).ToString()
            Dim middleTxt As New List(Of String)

            '最初マッチング分を切り取る
            'Dim IndexofMatching As Integer = FirstIndexOfMatching(currentTxt)
            'With currentTxt
            '    middleTxt.AddRange({ .Substring(0, IndexofMatching), .Substring(IndexofMatching, Text.Length), .Substring(IndexofMatching + Text.Length)})
            'End With

            'すべてマッチング分を切り取る
            Dim indexes() As Integer = IndexesOfMatching(currentTxt)
            With currentTxt
                middleTxt.Add(.Substring(0, indexes(0)))
                For i As Integer = 0 To indexes.Length - 1
                    middleTxt.Add(.Substring(indexes(i), Text.Length))
                    If Not i = (indexes.Length - 1) Then
                        middleTxt.Add(.Substring(indexes(i) + Text.Length, indexes(i + 1) - indexes(i) - Text.Length))
                    Else
                        middleTxt.Add(.Substring(indexes(i) + Text.Length))
                    End If
                Next
            End With

            'マッチング分をHighLightで表示
            Using g As Graphics = e.Graphics
                Dim rec As Rectangle = e.Bounds
                Dim foreColor As Color = Me.ForeColor
                Dim backColor As Color = Color.Transparent

                For Each str As String In middleTxt
                    rec.Width = TextRenderer.MeasureText(g, str, e.Font, New Size(Integer.MaxValue, Integer.MinValue), TextFormatFlags.NoPadding).Width
                    If IsSmContainsIgnoreCaseWidthaller(str) Then
                        foreColor = Color.Red
                        backColor = Color.LightYellow
                    Else
                        foreColor = Me.ForeColor
                        backColor = Color.Transparent
                    End If

                    TextRenderer.DrawText(g, str, e.Font, rec, foreColor, backColor, TextFormatFlags.NoPadding)
                    rec.Location = New Point(rec.X + rec.Width, rec.Y)

                Next
            End Using
        End If
    End Sub
End Class