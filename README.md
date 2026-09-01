Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long
    
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As String) As Long
    
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String) As Long
    
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const WM_SETTEXT = &HC
Private Const BM_CLICK = &HF5

Sub UnprotectWord2003Dot()
    Dim fd As Office.FileDialog
    Dim filePath As String
    Dim tempDoc As Document
    Dim targetDoc As Document
    
    ' ファイル選択
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Word 2003 Templates", "*.dot"
    
    If fd.Show = -1 Then
        filePath = fd.SelectedItems(1)
    Else
        Exit Sub
    End If
    
    ' Word 2003 では「テンプレートとして追加」で開くと保護が緩い
    ' 方法1: グローバルテンプレートとして読み込む
    Application.AddIns.Add filePath, False
    
    ' 方法2: テンプレートベースで新規文書作成
    Set tempDoc = Documents.Add(Template:=filePath, NewTemplate:=False)
    
    ' VBEを表示（パスワードダイアログが出る）
    On Error Resume Next
    Application.VBE.MainWindow.Visible = True
    
    ' パスワードダイアログが出たら自動入力
    Sleep 500
    BypassVBE2003Password
    
    ' マクロを抽出
    On Error GoTo ErrorHandler
    ExtractMacros2003 tempDoc
    
    ' クリーンアップ
    tempDoc.Close SaveChanges:=False
    
    MsgBox "解除完了。マクロは標準テンプレートにコピーされました。", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "エラー: " & Err.Description, vbExclamation
End Sub

Sub BypassVBE2003Password()
    Dim hwndVBE As Long
    Dim hwndDlg As Long
    Dim hwndEdit As Long
    Dim hwndBtn As Long
    Dim i As Integer
    
    ' VBEウィンドウを探す
    hwndVBE = FindWindow("wndclass_desked_gsk", vbNullString)
    
    ' プロジェクトエクスプローラで右クリック→プロパティを選択
    ' または直接パスワードダイアログを探す
    For i = 1 To 10
        Sleep 200
        
        ' 「プロジェクトのプロパティ」ダイアログ
        hwndDlg = FindWindow("#32770", "プロジェクトのプロパティ")
        If hwndDlg = 0 Then hwndDlg = FindWindow("#32770", "Project Properties")
        If hwndDlg = 0 Then hwndDlg = FindWindow("#32770", "VBAProject Properties")
        
        If hwndDlg <> 0 Then Exit For
    Next i
    
    If hwndDlg = 0 Then Exit Sub
    
    ' パスワード入力欄を探す（Word 2003ではEditコントロール）
    hwndEdit = FindWindowEx(hwndDlg, 0, "Edit", vbNullString)
    
    ' パスワードを入力（Word 2003では空文字でも通ることがある）
    If hwndEdit <> 0 Then
        SendMessage hwndEdit, WM_SETTEXT, 0, " "
    End If
    
    ' OKボタンを探してクリック
    hwndBtn = FindWindowEx(hwndDlg, 0, "Button", "OK")
    If hwndBtn <> 0 Then
        SendMessage hwndBtn, BM_CLICK, 0, 0
    End If
    
    Sleep 300
    
    ' エラーダイアログが出たら閉じる
    hwndDlg = FindWindow("#32770", "Microsoft Visual Basic")
    If hwndDlg <> 0 Then
        hwndBtn = FindWindowEx(hwndDlg, 0, "Button", "OK")
        If hwndBtn <> 0 Then SendMessage hwndBtn, BM_CLICK, 0, 0
    End If
End Sub

Sub ExtractMacros2003(doc As Document)
    Dim vbComp As Object
    Dim exportPath As String
    Dim i As Integer
    
    exportPath = Environ("TEMP") & "\VBA2003_Recovery\"
    If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath
    
    On Error Resume Next
    
    ' Word 2003 では VBComponents へのアクセス方法が異なる
    For Each vbComp In doc.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1 ' vbext_ct_StdModule
                vbComp.export exportPath & vbComp.Name & ".bas"
                Debug.Print "Exported: " & vbComp.Name & ".bas"
                
                ' 標準テンプレートにコピー
                NormalTemplate.VBProject.VBComponents.Import exportPath & vbComp.Name & ".bas"
                
            Case 2 ' vbext_ct_ClassModule
                vbComp.export exportPath & vbComp.Name & ".cls"
                NormalTemplate.VBProject.VBComponents.Import exportPath & vbComp.Name & ".cls"
                
            Case 3 ' vbext_ct_MSForm
                vbComp.export exportPath & vbComp.Name & ".frm"
                NormalTemplate.VBProject.VBComponents.Import exportPath & vbComp.Name & ".frm"
                
            Case 100 ' vbext_ct_Document
                ' ThisDocument など - コードを直接コピー
                Dim srcMod As Object, dstMod As Object
                Dim srcLines As String
                
                Set srcMod = vbComp.CodeModule
                If srcMod.CountOfLines > 0 Then
                    srcLines = srcMod.Lines(1, srcMod.CountOfLines)
                    
                    ' 標準テンプレートに新規モジュール作成
                    Set dstMod = NormalTemplate.VBProject.VBComponents.Add(1).CodeModule
                    dstMod.AddFromString srcLines
                End If
        End Select
    Next
    
    ' 標準テンプレートを保存
    NormalTemplate.Save
    
    MsgBox "マクロを標準テンプレートにコピーしました。" & vbCrLf & _
           "エクスポート先: " & exportPath, vbInformation
End Sub

' Word 2003 専用: Organizer を使った迂回
Sub OrganizerBypass2003()
    Dim fd As Office.FileDialog
    Dim filePath As String
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Filters.Clear
    fd.Filters.Add "Word 2003 Templates", "*.dot"
    
    If fd.Show = -1 Then filePath = fd.SelectedItems(1) Else Exit Sub
    
    ' Word 2003 では OrganizerDialog を直接開くと保護されていても操作可能
    Application.OrganizerCopy _
        Source:=filePath, _
        Destination:=NormalTemplate.FullName, _
        Name:="*", _
        Object:=wdOrganizerObjectProjectItems
    
    MsgBox "Organizer でコピーが完了しました。標準テンプレートを確認してください。", vbInformation
End Sub