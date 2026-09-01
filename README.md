Sub Quick2003Bypass()
    ' Word 2003 では Documents.Open の引数に特定の組み合わせを使うと保護が無視される
    Dim d As Document
    Set d = Documents.Open("C:\path\to\file.dot", , , , "dummy", , , , , , True)
    Application.VBE.MainWindow.Visible = True
End Sub