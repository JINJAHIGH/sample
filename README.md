Sub ExtractMacros()
    Dim doc As Document
    Dim vbProj As Object
    Dim comp As Object
    
    ' パスワード保護されたファイルを開く（パスワード入力が必要）
    Set doc = Documents.Open("C:\path\to\template.dot", PasswordDocument:="")
    
    ' モジュールをテキストファイルとしてエクスポート
    For Each comp In doc.VBProject.VBComponents
        If comp.Type = 1 Then ' 標準モジュール
            comp.Export "C:\temp\" & comp.Name & ".bas"
        End If
    Next
    
    doc.Close SaveChanges:=False
    
    ' 新規テンプレートにインポート
    Documents.Add Template:="Normal", NewTemplate:=True
    For Each comp In ActiveDocument.VBProject.VBComponents
        If comp.Type = 1 Then
            ActiveDocument.VBProject.VBComponents.Import "C:\temp\" & comp.Name & ".bas"
        End If
    Next
    
    ActiveDocument.SaveAs "C:\path\to\new_template.dot"
End Sub