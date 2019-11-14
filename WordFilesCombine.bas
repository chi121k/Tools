Attribute VB_Name = "Module1"
Sub Wordファイル結合処理()

    Const delStringCnt As Integer = 9

    'フォルダの選択
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "フォルダを選択"
        .AllowMultiSelect = False
    
        If .Show = -1 Then
            mypath = .SelectedItems(1) & ""
        Else
            Exit Sub
        End If
    End With
    
    Dim doc_name As String
    ChDir mypath
    doc_name = Dir("*基本仕様書*.docx")
    
    '記録用ファイル作成
    Documents.Add
    
    Dim sec_cnt As Integer
    sec_cnt = 1
    Do While doc_name <> ""
    
        With Documents.Open(FileName:=mypath & "\" & doc_name, Visible:=False, ReadOnly:=True)
            .Content.Copy
            
            '読み取り専用ファイルを開いた場合の対処（そのまま閉じる）
            If .ReadOnly = True Then
                .Close SaveChanges:=wdDoNotSaveChanges
            Else
                .Close
            End If
        End With
        
        'ファイル内容のコピペ処理
        With Selection
            .EndKey Unit:=wdStory
            '.InsertFile doc_name
            .PasteAndFormat (wdFormatOriginalFormatting)
            .InsertBreak wdSectionBreakNextPage
            
            .Collapse wdCollapseEnd
        End With
        
        'ヘッダー設定
        ActiveDocument.Sections(sec_cnt).Headers(wdHeaderFooterPrimary).LinkToPrevious = False
        
        Set headerRange = ActiveDocument.Sections(sec_cnt). _
            Headers(wdHeaderFooterPrimary).Range
        
        With headerRange
        
            'ファイル名の挿入（テキスト）
            .Text = Replace(Replace(doc_name, Left(doc_name, delStringCnt), ""), ".docx", "")
            
            '中央揃え
            .Paragraphs.Alignment = wdAlignParagraphCenter
        
        End With
        
        Set headerRange = Nothing
        
        'フッター設定
        With ActiveDocument.Sections(sec_cnt)
            .Footers(wdHeaderFooterPrimary).PageNumbers.Add _
                PageNumberAlignment:=wdAlignPageNumberCenter
        End With
        
        doc_name = Dir
        sec_cnt = sec_cnt + 1
    Loop

End Sub
