Imports System.Windows.Forms
Imports wd = Microsoft.Office.Interop.Word '引用別名（Alias） https://social.technet.microsoft.com/wiki/contents/articles/32449.namespace-aliases-in-visual-basic-net-howto.aspx
Imports System.Runtime.InteropServices

Module Module2
    Sub Main_Module2()
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
        Dim app As New wd.Application
        Dim docs As wd.Documents = app.Documents
        Dim doc As wd.Document = docs.Open(fName,, [ReadOnly]:=True, AddToRecentFiles:=False)
        '不想把開啟的檔案加入最近開過清單
        Dim docContent As String = doc.Content.Text
        Dim startPst As Long = InStr(docContent, findTxt)
        If startPst > 0 Then
            Dim thisparaEnd As Long = InStr(startPst, docContent, Chr(13))
            Dim thisparaStart As Long = InStrRev(docContent, Chr(13), startPst) + 1
            If thisparaStart = 0 Then thisparaStart = 1 '如果是第一段文件
            Dim findTxtPara As String = Mid(docContent, thisparaStart, thisparaEnd - thisparaStart + 1)
            '以chr(13)段落標記來找前後分段處
            '    Dim foundRng As wd.Range = doc.Range()
            'Dim p As wd.Paragraph
            Console.WriteLine(findTxtPara)
            Console.ReadLine() '在Console按下Enter鍵即可離開
        End If
        '如果有用到Dim app As New wd.Application 這行，就最好執行此行:
        app.Quit(wd.WdSaveOptions.wdDoNotSaveChanges)
        doc = Nothing : docs = Nothing : app = Nothing
    End Sub
End Module
