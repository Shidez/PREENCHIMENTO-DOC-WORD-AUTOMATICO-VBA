Attribute VB_Name = "Módulo2"
Public Sub gerarDocumentos()

    Dim docWord As Word.Document 'variavel do novo word
    Dim wordApp As Word.Application 'variavel do app word
    
    Dim tb As ListObject 'tabela TB_Pessoas
    Dim tr As ListRow 'linha da tabela
   
    On Error Resume Next
   
    Set wordApp = New Word.Application
    'wordApp.Visible = True
   
    'Planilha1 é o nome da aba da planilha
    'TB_PESSOAS é o nome da planilha que vc colocou interno no Excel
    
    Set tb = Planilha1.ListObjects("TB_PESSOAS")
       
    For Each tr In tb.ListRows
        'pegar template na pasta abaixo:
        Set docWord = wordApp.Documents.Open("caminho completo pasta para salvar")
        'preencher os marcadores do template
        docWord.Bookmarks("awb").Range.Text = tr.Range(, 1)
        'docWord.Bookmarks("data2").Range.Text = tr.Range(, 2)
        docWord.Bookmarks("analista").Range.Text = tr.Range(, 2)
        docWord.Bookmarks("profissao").Range.Text = tr.Range(, 3)
        docWord.SaveAs2 docWord.Path & "\" & tr.Range(, 1) & ".docx"
        
        'se erro de linha vazia:
        If Err = 4198 Then
            docWord.Close SaveChanges:=False
            wordApp.Quit
            MsgBox "Documentos com linhas preenchidas gerados com sucesso!"
            Exit Sub
            
        'se não, salvar novo documento
        Else
        docWord.Close SaveChanges:=True
        End If
        
    Next

wordApp.Quit

MsgBox "Documentos gerados com sucesso!"

End Sub

