Attribute VB_Name = "Module1"
Sub Word�t�@�C����������()

    Const delStringCnt As Integer = 9

    '�t�H���_�̑I��
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "�t�H���_��I��"
        .AllowMultiSelect = False
    
        If .Show = -1 Then
            mypath = .SelectedItems(1) & ""
        Else
            Exit Sub
        End If
    End With
    
    Dim doc_name As String
    ChDir mypath
    doc_name = Dir("*��{�d�l��*.docx")
    
    '�L�^�p�t�@�C���쐬
    Documents.Add
    
    Dim sec_cnt As Integer
    sec_cnt = 1
    Do While doc_name <> ""
    
        With Documents.Open(FileName:=mypath & "\" & doc_name, Visible:=False, ReadOnly:=True)
            .Content.Copy
            
            '�ǂݎ���p�t�@�C�����J�����ꍇ�̑Ώ��i���̂܂ܕ���j
            If .ReadOnly = True Then
                .Close SaveChanges:=wdDoNotSaveChanges
            Else
                .Close
            End If
        End With
        
        '�t�@�C�����e�̃R�s�y����
        With Selection
            .EndKey Unit:=wdStory
            '.InsertFile doc_name
            .PasteAndFormat (wdFormatOriginalFormatting)
            .InsertBreak wdSectionBreakNextPage
            
            .Collapse wdCollapseEnd
        End With
        
        '�w�b�_�[�ݒ�
        ActiveDocument.Sections(sec_cnt).Headers(wdHeaderFooterPrimary).LinkToPrevious = False
        
        Set headerRange = ActiveDocument.Sections(sec_cnt). _
            Headers(wdHeaderFooterPrimary).Range
        
        With headerRange
        
            '�t�@�C�����̑}���i�e�L�X�g�j
            .Text = Replace(Replace(doc_name, Left(doc_name, delStringCnt), ""), ".docx", "")
            
            '��������
            .Paragraphs.Alignment = wdAlignParagraphCenter
        
        End With
        
        Set headerRange = Nothing
        
        '�t�b�^�[�ݒ�
        With ActiveDocument.Sections(sec_cnt)
            .Footers(wdHeaderFooterPrimary).PageNumbers.Add _
                PageNumberAlignment:=wdAlignPageNumberCenter
        End With
        
        doc_name = Dir
        sec_cnt = sec_cnt + 1
    Loop

End Sub
