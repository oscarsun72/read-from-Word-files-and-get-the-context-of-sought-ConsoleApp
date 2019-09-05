Imports System.Windows.Forms
Imports wd = Microsoft.Office.Interop.Word '引用別名（Alias） https://social.technet.microsoft.com/wiki/contents/articles/32449.namespace-aliases-in-visual-basic-net-howto.aspx

Module Module1
    Sub Main()
        Main_Module2()
    End Sub
    Sub Main_original()
        Dim fName = CurDir() & "\心經.docx"
        Dim findTxt = "色即是空"
        Dim f As New Form 'https://social.msdn.microsoft.com/Forums/windowsdesktop/zh-TW/46d73fc9-4603-43ae-acf8-03873f17dfeb/msgbox-topmost?forum=232
        f.TopMost = True '訊息方塊最上層顯示
        If Dir(fName) = "" Then
            MessageBox.Show(f, "沒有此檔案，請檢查路徑、檔名是否正確！") '訊息方塊最上層顯示,下式不行：
            'MsgBox("沒有此檔案，請檢查路徑、檔名是否正確！", Buttons:=MsgBoxStyle.Critical + MsgBoxStyle.ApplicationModal)
            Exit Sub
        End If
        If findTxt = "" Then
            MessageBox.Show(f, "沒有尋找字串，請重新指定！")
            'MsgBox("沒有尋找字串，請重新指定！",Buttons:=MsgBoxStyle.Critical + MsgBoxStyle.ApplicationModal)
            Exit Sub
        End If
        Dim app As New wd.Application
        Dim docs As wd.Documents
        Dim doc As wd.Document
        docs = app.Documents 'Use the Documents property to return the Documents collection.
        'And use the return Documents collection to initialize the docs Documents
        'https://docs.microsoft.com/zh-tw/dotnet/api/microsoft.office.interop.word.documents.add?view=word-pia#Microsoft_Office_Interop_Word_Documents_Add_System_Object__System_Object__System_Object__System_Object__
        app.Visible = True '是怕開啟檔案時會有對話方塊，如果當機，才能手動關閉Word app
        app.WindowState = wd.WdWindowState.wdWindowStateMinimize
        doc = docs.Open(fName) '如果啟動時有對話方塊，須手動關閉程式才能繼續
        'doc = GetObject(fName) '現有的檔案用GetObject即可，就不必以上'Dim app As New wd.Application以下的程式碼了，直接用這一行取代即可
        '可見用GetObject的話，程式碼精簡許多
        Dim foundRng As wd.Range = doc.Range()
        If foundRng.Find().Execute(findTxt) Then '如果有找到的話
            foundRng.Select()
            Console.WriteLine(doc.ActiveWindow.Selection.Paragraphs(1).Range.Text)
            'WriteLine和Write不同在於WriteLine會將插入點置於印出來的文字的下一行，就不會與Ctrl+F5執行後產生的提示文字重疊
            Console.ReadLine() '在Console按下Enter鍵即可離開
        End If
        doc.ActiveWindow.Visible = True
        'doc.Close(wd.WdSaveOptions.wdDoNotSaveChanges) '如果要關掉文件，再執行此行
        '如果有用到Dim app As New wd.Application 這行，就最好執行此行:
        app.Quit(wd.WdSaveOptions.wdDoNotSaveChanges)
        doc = Nothing : docs = Nothing : app = Nothing
    End Sub
End Module
