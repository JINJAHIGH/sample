Sub RecoverFromProtectedDot()
    Dim protectedFile As String
    protectedFile = "C:\OldFiles\template.dot"
    
    ' 保護されたファイルから Normal.dot へ強制コピー
    Application.OrganizerCopy _
        Source:=protectedFile, _
        Destination:=NormalTemplate.FullName, _
        Name:="*", _
        Object:=wdOrganizerObjectProjectItems
    
    ' Normal.dot を保存
    NormalTemplate.Save
    
    ' これで Normal.dot からマクロをエクスポートできる
    MsgBox "マクロを Normal.dot にコピーしました。" & vbCrLf & _
           "VBAエディタで確認してください。", vbInformation
End Sub