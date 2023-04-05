Option Explicit
Dim keyword, shell
If Not IsTextSelected Then
  SelectWord()
End If
keyword = GetSelectedString()
' 選択されている文字列がある場合のみGoogle検索'
If keyword <> "" Then
  Set shell = CreateObject("Shell.Application")
  shell.ShellExecute "https://www.google.co.jp/search?q=" & keyword
End If