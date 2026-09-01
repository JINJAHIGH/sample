Sub BreakPassword()
    Dim i As Integer, FileName As String
    FileName = "C:\path\to\your\file.dot"
    
    ' 保護されたファイルを開く
    Documents.Open FileName, PasswordDocument:="", _
        WritePasswordDocument:="", Revert:=False
    
    ' マクロモジュールをエクスポート・インポートで再設定
    Application.OrganizerCopy Source:=FileName, _
        Destination:=NormalTemplate, Name:="Module1", _
        Object:=wdOrganizerObjectProjectItems
End Sub