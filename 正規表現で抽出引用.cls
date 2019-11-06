Function GetStr(rng As Range, Patt)
  With CreateObject("VBscript.regexp")
    .Global = True
    .Pattern = Patt
    If .Execute(rng).Count = 0 Then
      GetStr = ""
    Else
      GetStr = .Execute(rng)(0)
    End If
  End With
End Function

##########################################3
#導入
#（Windows）Alt+F11でVBA開く→挿入→標準モジュール→コードを貼り付け→保存
#
#使い方
#=GetStr(セル名,パターン)
#例
#=GetStr(A4,"(F0|W0)\d{1,9}")
#（F0またはW0から始まって、1桁～9桁の数字で終わる文字列を抽出）
#
#正規表現の参照
#https://murashun.jp/blog/20190215-01.html

