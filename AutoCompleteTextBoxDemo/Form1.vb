Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim nameArray() As String = {"ALualuAuAuu", "al", "ＡＬ", "チェック", "ちぇっく", "ﾁｪｯｸ", "1234567890ABCDEFGﾁｪｯｸ"}
        txtFeature.AutoCompleteValues = nameArray
    End Sub
End Class
