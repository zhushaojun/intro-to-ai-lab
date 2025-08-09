Option Explicit

Sub BuildZhAcademicTemplateAndNumbering_Compat()
    Dim doc As Document
    Dim s As Style
    Dim ftr As HeaderFooter
    Dim lt As ListTemplate

    Set doc = ActiveDocument

    ' 页面设置 A4 + 边距
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

    ' 正文 Normal：宋体 小四 1.5 倍行距 首行缩进约 2 字符（按 1 字符 ≈ 0.74cm 近似）
    With doc.Styles(wdStyleNormal).Font
        .NameFarEast = "宋体"
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

    ' 标题1：黑体 三号，加粗，单倍，段前12 段后6，与下段同页
    Set s = doc.Styles(wdStyleHeading1)
    With s.Font
        .NameFarEast = "黑体"
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

    ' 标题2：黑体 小三
    Set s = doc.Styles(wdStyleHeading2)
    With s.Font
        .NameFarEast = "黑体"
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

    ' 标题3：黑体 四号
    Set s = doc.Styles(wdStyleHeading3)
    With s.Font
        .NameFarEast = "黑体"
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

    ' 代码：Code 样式（浅灰底，左边框）
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

    ' 代码：Code Block 样式（Pandoc 兼容名）
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

    ' 图表题：Caption 样式
    Set s = doc.Styles(wdStyleCaption)
    With s.Font
        .NameFarEast = "宋体"
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

    ' 列表：List Paragraph（左缩进约 1 字符）
    Set s = doc.Styles(wdStyleListParagraph)
    With s.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0.74)
        .FirstLineIndent = 0
        .LineSpacingRule = wdLineSpace1pt5
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With

    ' 参考文献：Bibliography（悬挂缩进约 2 字符）
    On Error Resume Next
    Set s = doc.Styles("Bibliography")
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="Bibliography", Type:=wdStyleTypeParagraph)
    On Error GoTo 0
    With s.Font
        .NameFarEast = "宋体"
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

    ' 页码：页脚居中
    Set ftr = doc.Sections(1).Footers(wdHeaderFooterPrimary)
    ftr.Range.Text = ""
    ftr.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    ftr.Range.Fields.Add Range:=ftr.Range, Type:=wdFieldPage

    ' 多级自动编号：标题1/2/3 → 1、1.1、1.1.1
    Set lt = doc.ListTemplates.Add(OutlineNumbered:=True, Name:="CN-标题s")

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
        .LinkedStyle = "标题 1"
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
        .LinkedStyle = "标题 2"
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
        .LinkedStyle = "标题 3"
    End With

    doc.Styles(wdStyleHeading1).LinkToListTemplate lt, 1
    doc.Styles(wdStyleHeading2).LinkToListTemplate lt, 2
    doc.Styles(wdStyleHeading3).LinkToListTemplate lt, 3

    ' 保存为模板（docx 参考模板）——兼容 Word（无 GetSaveAsFilename）
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
        basePath = "D:\QSync\work\教学\人工智能导论实验指导书"
        If Dir(basePath, vbDirectory) = "" Then
            basePath = Environ$("USERPROFILE") & "\Desktop"
        End If
        If Right$(basePath, 1) <> "\\" Then basePath = basePath & "\\"
        savePath = basePath & "reference-zh-academic.docx"
    End If

    doc.SaveAs2 savePath, wdFormatXMLDocument
    MsgBox "已保存模板：" & savePath, vbInformation
End Sub

