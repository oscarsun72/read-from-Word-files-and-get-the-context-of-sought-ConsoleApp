Imports System.Windows.Forms
Imports wd = Microsoft.Office.Interop.Word '引用別名（Alias） https://social.technet.microsoft.com/wiki/contents/articles/32449.namespace-aliases-in-visual-basic-net-howto.aspx
Imports System.Runtime.InteropServices

Module Module2
    Sub Main_Module2()
        Dim fName = CurDir() & "\翁方綱及其文獻學研究_print.doc" '  "\心經.docx" 
        Dim findTxt = "我愛，與我母。" ' "色即是空" '
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
        Dim doc As wd.Document = GetObject(fName)
        Dim foundRng As wd.Range = doc.Range()
        If foundRng.Find().Execute(findTxt) Then '如果有找到的話
            foundRng.Select()
            Console.WriteLine(doc.ActiveWindow.Selection.Paragraphs(1).Range.Text)
            'WriteLine和Write不同在於WriteLine會將插入點置於印出來的文字的下一行，就不會與Ctrl+F5執行後產生的提示文字重疊
            Console.ReadLine() '在Console按下Enter鍵即可離開
        End If
        'doc.ActiveWindow.Visible = True
        'doc.Close(wd.WdSaveOptions.wdDoNotSaveChanges) '如果要關掉文件，再執行此行
        '如果有用到Dim app As New wd.Application 這行，就最好執行此行:
        doc.Application.Quit(wd.WdSaveOptions.wdDoNotSaveChanges)
        doc = Nothing
    End Sub
End Module
