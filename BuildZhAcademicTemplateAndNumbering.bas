Option Explicit

Sub BuildZhAcademicTemplateAndNumbering_Compat()
    Dim doc As Document
    Dim s As Style
    Dim ftr As HeaderFooter
    Dim lt As ListTemplate

    Set doc = ActiveDocument

    ' ҳ������ A4 + �߾�
    With doc.PageSetup
        .PageWidth = CentimetersToPoints(21#)
        .PageHeight = CentimetersToPoints(29.7)
        .TopMargin = CentimetersToPoints(2.5)
        .BottomMargin = CentimetersToPoints(2.5)
        .LeftMargin = CentimetersToPoints(3#)
        .RightMargin = CentimetersToPoints(2.5)
        .HeaderDistance = CentimetersToPoints(1.5)
        .FooterDistance = CentimetersToPoints(1.5)
    End With

    ' ���� Normal������ С�� 1.5 ���о� ��������Լ 2 �ַ����� 1 �ַ� �� 0.74cm ���ƣ�
    With doc.Styles(wdStyleNormal).Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .Size = 12
    End With
    With doc.Styles(wdStyleNormal).ParagraphFormat
        .LineSpacingRule = wdLineSpace1pt5
        .FirstLineIndent = CentimetersToPoints(1.48)
        .LeftIndent = 0
        .RightIndent = 0
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With

    ' ����1������ ���ţ��Ӵ֣���������ǰ12 �κ�6�����¶�ͬҳ
    Set s = doc.Styles(wdStyleHeading1)
    With s.Font
        .NameFarEast = "����"
        .NameAscii = "Arial"
        .Bold = True
        .Size = 16
    End With
    With s.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .SpaceBefore = 12
        .SpaceAfter = 6
        .KeepWithNext = True
    End With

    ' ����2������ С��
    Set s = doc.Styles(wdStyleHeading2)
    With s.Font
        .NameFarEast = "����"
        .NameAscii = "Arial"
        .Bold = True
        .Size = 15
    End With
    With s.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .SpaceBefore = 12
        .SpaceAfter = 6
        .KeepWithNext = True
    End With

    ' ����3������ �ĺ�
    Set s = doc.Styles(wdStyleHeading3)
    With s.Font
        .NameFarEast = "����"
        .NameAscii = "Arial"
        .Bold = True
        .Size = 14
    End With
    With s.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .SpaceBefore = 12
        .SpaceAfter = 6
        .KeepWithNext = True
    End With

    ' ���룺Code ��ʽ��ǳ�ҵף���߿�
    On Error Resume Next
    Set s = doc.Styles("Code")
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="Code", Type:=wdStyleTypeParagraph)
    On Error GoTo 0
    With s.Font
        .Name = "Consolas"
        .Size = 10
    End With
    With s.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .LeftIndent = CentimetersToPoints(0.74)
        .SpaceBefore = 6
        .SpaceAfter = 6
        .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
        .Borders(wdBorderLeft).Color = wdColorGray25
    End With
    With s.ParagraphFormat.Shading
        .BackgroundPatternColorIndex = wdGray25
    End With

    ' ���룺Code Block ��ʽ��Pandoc ��������
    On Error Resume Next
    Set s = doc.Styles("Code Block")
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="Code Block", Type:=wdStyleTypeParagraph)
    On Error GoTo 0
    With s.Font
        .Name = "Consolas"
        .Size = 10
    End With
    With s.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .LeftIndent = CentimetersToPoints(0.74)
        .SpaceBefore = 6
        .SpaceAfter = 6
        .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
        .Borders(wdBorderLeft).Color = wdColorGray25
    End With
    With s.ParagraphFormat.Shading
        .BackgroundPatternColorIndex = wdGray25
    End With

    ' ͼ���⣺Caption ��ʽ
    Set s = doc.Styles(wdStyleCaption)
    With s.Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .Size = 10.5
        .Bold = False
    End With
    With s.ParagraphFormat
        .Alignment = wdAlignParagraphCenter
        .LineSpacingRule = wdLineSpaceSingle
        .SpaceBefore = 6
        .SpaceAfter = 6
    End With

    ' �б�List Paragraph��������Լ 1 �ַ���
    Set s = doc.Styles(wdStyleListParagraph)
    With s.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0.74)
        .FirstLineIndent = 0
        .LineSpacingRule = wdLineSpace1pt5
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With

    ' �ο����ף�Bibliography����������Լ 2 �ַ���
    On Error Resume Next
    Set s = doc.Styles("Bibliography")
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="Bibliography", Type:=wdStyleTypeParagraph)
    On Error GoTo 0
    With s.Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .Size = 12
    End With
    With s.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .LeftIndent = CentimetersToPoints(1.48)
        .FirstLineIndent = CentimetersToPoints(-1.48)
        .SpaceBefore = 0
        .SpaceAfter = 6
    End With

    ' ҳ�룺ҳ�ž���
    Set ftr = doc.Sections(1).Footers(wdHeaderFooterPrimary)
    ftr.Range.Text = ""
    ftr.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    ftr.Range.Fields.Add Range:=ftr.Range, Type:=wdFieldPage

    ' �༶�Զ���ţ�����1/2/3 �� 1��1.1��1.1.1
    Set lt = doc.ListTemplates.Add(OutlineNumbered:=True, Name:="CN-����s")

    With lt.ListLevels(1)
        .NumberFormat = "%1"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0.74)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        .LinkedStyle = "���� 1"
    End With

    With lt.ListLevels(2)
        .NumberFormat = "%1.%2"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.48)
        .TabPosition = wdUndefined
        .ResetOnHigher = 1
        .StartAt = 1
        .LinkedStyle = "���� 2"
    End With

    With lt.ListLevels(3)
        .NumberFormat = "%1.%2.%3"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(2.22)
        .TabPosition = wdUndefined
        .ResetOnHigher = 2
        .StartAt = 1
        .LinkedStyle = "���� 3"
    End With

    doc.Styles(wdStyleHeading1).LinkToListTemplate lt, 1
    doc.Styles(wdStyleHeading2).LinkToListTemplate lt, 2
    doc.Styles(wdStyleHeading3).LinkToListTemplate lt, 3

    ' ����Ϊģ�壨docx �ο�ģ�壩�������� Word���� GetSaveAsFilename��
    Dim savePath As String
    Dim fd As Object ' late binding to avoid Office library reference

    On Error Resume Next
    Set fd = Application.FileDialog(2) ' msoFileDialogSaveAs = 2
    On Error GoTo 0

    If Not fd Is Nothing Then
        With fd
            .InitialFileName = Environ$("USERPROFILE") & "\Desktop\reference-zh-academic.docx"
            If .Show = -1 Then
                savePath = .SelectedItems(1)
            End If
        End With
    End If

    If Len(savePath) = 0 Then
        Dim basePath As String
        basePath = "D:\QSync\work\��ѧ\�˹����ܵ���ʵ��ָ����"
        If Dir(basePath, vbDirectory) = "" Then
            basePath = Environ$("USERPROFILE") & "\Desktop"
        End If
        If Right$(basePath, 1) <> "\\" Then basePath = basePath & "\\"
        savePath = basePath & "reference-zh-academic.docx"
    End If

    doc.SaveAs2 savePath, wdFormatXMLDocument
    MsgBox "�ѱ���ģ�壺" & savePath, vbInformation
End Sub

