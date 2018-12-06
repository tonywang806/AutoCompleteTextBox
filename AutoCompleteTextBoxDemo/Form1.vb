Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim nameArray() As String = {"11111111", "あああああ111", "xx11aa", "11111111ああああああああああああ", "A1111Lual111111uAuAuu", "al", "ＡＬ", "チェック", "ちぇ っく", "ﾁｪｯｸ", "123456  7890ABCDEFGﾁｪｯｸ", "ＰＢ", "pb", "PB", "Pb", "pB", "ｐｂ", "Ｐｂ", "ｐＢ", "PＢ", "Pｂ", "pＢ", "pｂ", "ＰB", "Ｐb", "ｐB", "ｐb"}
        txtFeature.AutoCompleteValues = nameArray
    End Sub
End Class
