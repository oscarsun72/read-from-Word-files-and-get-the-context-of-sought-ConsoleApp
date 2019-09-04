Imports System.Windows.Forms
Imports wd = Microsoft.Office.Interop.Word '引用別名（Alias） https://social.technet.microsoft.com/wiki/contents/articles/32449.namespace-aliases-in-visual-basic-net-howto.aspx
Imports System.Runtime.InteropServices

Module Module2_Marshal
    Sub Main_Marshal()
        Dim fName = CurDir() & "\翁方綱及其文獻學研究_print.doc" '"\心經.docx"
        Dim findTxt = "我愛，與我母。"
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
        Dim obj As [Object] = Nothing 'https://docs.microsoft.com/zh-tw/dotnet/api/system.runtime.interopservices.marshal.getactiveobject?view=netframework-4.8
        Try
            obj = Marshal.GetActiveObject("Word.Application")
        Catch e As Exception
        End Try
        If obj Is Nothing Then
            Dim app As New wd.Application '只是把app改成obj而已，似乎多此一舉，然若能在出錯的版本上執行，也算權解了
            obj = Marshal.GetActiveObject("Word.Application")
        End If
        Dim docs As wd.Documents = obj.Documents
        'Use the Documents property to return the Documents collection.
        'And use the return Documents collection to initialize the docs Documents
        'https://docs.microsoft.com/zh-tw/dotnet/api/microsoft.office.interop.word.documents.add?view=word-pia#Microsoft_Office_Interop_Word_Documents_Add_System_Object__System_Object__System_Object__System_Object__
        Dim doc As wd.Document = docs.Open(fName)
        'ojc.Visible = True '是怕開啟檔案時會有對話方塊，如果當機，才能手動關閉Word app
        'app.WindowState = wd.WdWindowState.wdWindowStateMinimize
        'Dim foundRng As wd.Range = doc.Range()
        'Dim sel As wd.Selection = doc.ActiveWindow.Selection
        Dim p As wd.Paragraph
        For Each p In doc.Paragraphs
            If InStr(p.Range.Text, findTxt) > 0 Then '如果有找到的話
                Console.WriteLine(p.Range.Text)
                Exit For
            End If
        Next
        'If sel.Find().Execute(findTxt) = True Then '如果有找到的話
        'If foundRng.Find().Execute(findTxt) = True Then '如果有找到的話
        'foundRng.Select()
        'Console.WriteLine(doc.ActiveWindow.Selection.Paragraphs(1).Range.Text)
        'Console.WriteLine(sel.Paragraphs(1).Range.Text)
        'WriteLine和Write不同在於WriteLine會將插入點置於印出來的文字的下一行，就不會與Ctrl+F5執行後產生的提示文字重疊
        Console.ReadLine() '在Console按下Enter鍵即可離開
        'End If
        'doc.ActiveWindow.Visible = True
        'doc.Close(wd.WdSaveOptions.wdDoNotSaveChanges) '如果要關掉文件，再執行此行
        '如果有用到Dim app As New wd.Application 這行，就最好執行此行:
        obj.Quit(wd.WdSaveOptions.wdDoNotSaveChanges)
        doc = Nothing : docs = Nothing : obj = Nothing
    End Sub
End Module
