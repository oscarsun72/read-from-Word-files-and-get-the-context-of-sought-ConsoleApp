Imports Microsoft.Office.Interop.Word
Module Module1
    Sub Main()
        Const fName = "c:\心經.docx"
        Dim findTxt = "恐怖"
        Dim doc As Microsoft.Office.Interop.Word.Document
        'Dim docs As Microsoft.Office.Interop.Word.Documents
        doc = GetObject(fName) '現有的檔案用GetObject即可
        Dim foundRng As Microsoft.Office.Interop.Word.Range = doc.Range()
        If foundRng.Find().Execute(findTxt) Then '如果有找到的話
            foundRng.Select()
            Console.Write(doc.ActiveWindow.Selection.Paragraphs(1).Range.Text)
            Console.ReadLine()
        End If
        doc.ActiveWindow.Visible = True
        'doc.Close(WdSaveOptions.wdDoNotSaveChanges)'如果要關掉文件，再執行此行
        doc = Nothing
    End Sub
End Module
