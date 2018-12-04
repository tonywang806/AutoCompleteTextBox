Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim nameArray() As String = {"ALualuAuAuu", "al", "ＡＬ", "チェック", "ちぇっく", "ﾁｪｯｸ", "1234567890ABCDEFGﾁｪｯｸ", "ＰＢ", "pb", "PB", "Pb", "pB", "ｐｂ", "Ｐｂ", "ｐＢ", "PＢ", "Pｂ", "pＢ", "pｂ", "ＰB", "Ｐb", "ｐB", "ｐb"}
        txtFeature.AutoCompleteValues = nameArray
    End Sub
End Class
