Attribute VB_Name = "MarkupMacrosesXP"
Dim NewDoc$
Dim CurDoc$

Sub ReturnCellContentsToArray()
    Dim intCells As Integer
    Dim celTable As Cell
    Dim strCells() As String
    Dim intCount As Integer
    Dim rngText As Range

    If ActiveDocument.Tables.Count >= 1 Then
        With ActiveDocument.Tables(1).Range
            intCells = .Cells.Count
            ReDim strCells(intCells)
            intCount = 1
            For Each celTable In .Cells
                Set rngText = celTable.Range
                rngText.MoveEnd Unit:=wdCharacter, Count:=-1
                strCells(intCount) = rngText
                intCount = intCount + 1
            Next celTable
        End With
    End If
End Sub



Sub VenturaPrepTableXP()
Attribute VenturaPrepTableXP.VB_Description = "Macro recorded 10.10.2006 by Artem"
Attribute VenturaPrepTableXP.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro4"
'
' Macro4 Macro
' Macro recorded 10.10.2006 by Artem
'
Dim t As Table
Dim C As Cell
Dim R As Row
Dim Maxi As Long

ActiveDocument.ActiveWindow.View.ShowAll = True
For Each t In ActiveDocument.Tables
    Maxi = t.Columns.Count
    t.PreferredWidthType = wdPreferredWidthPercent
    t.PreferredWidth = 100
    t.Columns.PreferredWidthType = _
        wdPreferredWidthPercent
    t.Columns.PreferredWidth = 100 / Maxi
    'На случай если сука-ворд не обновляет формат таблицы
    t.Select
    Selection.Cut
    Selection.Paste
    Set t = Selection.Tables(1)
    t.PreferredWidth = 100 'Добавлено ПА 24.07.2007 (см. предыдущ. коммент)
    '----------------------------------------------------
    t.Cell(1, 1).Select
    Selection.Collapse wdCollapseStart
    Selection.InsertRowsAbove 1
    Selection.Cells.Split NumRows:=1, NumColumns:=Maxi, MergeBeforeSplit:=True
    'T.Cell(T.Rows.Count, 1).Select 'Ошибка на этой стадии означает, что надо вставить 1 доп строку внизу таблицы
    t.Rows.Add
    t.Select: Selection.EndKey: Selection.SelectRow
    Selection.Cells.Split NumRows:=1, NumColumns:=Maxi, MergeBeforeSplit:=True
    For Each C In t.Range.Cells
        C.Select
        With Selection
            If Not .Text = Chr(13) & Chr(7) Then
                .Collapse: .SelectCell 'на случай объединенных ячеек
                .MoveEnd wdCharacter, -1
                .Cut
                .Font.Reset
                .TypeText "@Z_TBL_CELL_BEG = "
                .TypeParagraph
                .Collapse wdCollapseEnd
                .Paste
                .Font.Reset
                .TypeParagraph: .TypeText "@Z_TBL_CELL_END = "
            Else
                .Font.Reset
                .TypeText "@Z_TBL_CELL_BEG = "
                .TypeParagraph
                .TypeText "@Z_TBL_CELL_END = "
            End If

        End With
    Next C
    'Проверка вертикальных объединений
    For Each C In t.Range.Cells
        C.Select
        With Selection
            i = .Information(wdEndOfRangeRowNumber) - _
                 .Information(wdStartOfRangeRowNumber)
            If i > 0 Then
                '.SelectColumn: lngClmn = .Cells(1).ColumnIndex
                
                C.Split i + 1
'                .MoveDown Unit:=wdLine, Count:=1
'                .MoveUp Unit:=wdLine, Count:=1
'                For i = i To 1 Step -1
'                    .Font.Reset
'                    .TypeText "@Z_TBL_CELL_BEG = VJOINED"
'                    .TypeParagraph: .TypeText "@Z_TBL_CELL_END = "
'                    .MoveDown Unit:=wdLine, Count:=1
'                Next i
            End If
            If .Text = Chr(13) & Chr(7) Then
                    .Font.Reset
                    .TypeText "@Z_TBL_CELL_BEG = VJOINED"
                    .TypeParagraph: .TypeText "@Z_TBL_CELL_END = "
            End If

        End With
    Next C
    'Проверка горизонтальных объединений
   t.PreferredWidth = 100 'Добавлено ПА 24.07.2007 (сука Ворд...)

    For Each R In t.Rows
        'i = 1
        'ii = 1
        R.Cells(1).Select: Selection.Collapse wdCollapseStart
        If Not R.Index = 2 Then Selection.TypeText "@Z_TBL_ROW_END = ": Selection.TypeParagraph
        Selection.TypeText "@Z_TBL_ROW_BEG = ": Selection.TypeParagraph
        If R.Cells.Count < Maxi Then
            SplitHorCells R, Maxi
        End If
    Next R
   With Selection
   t.Cell(t.Rows.Count, 1).Select: .Rows.Delete
   t.Cell(1, 1).Select: .Rows.Delete
    
    .Collapse wdCollapseStart
    '.TypeText "@Z_STYLE70 = ": .TypeParagraph Отключено ПА 19.07.2007 (маркер версии ставится в Markup и SplitDocTotxt
    .TypeText "@Z_TBL_BEG = VERSION(10), TAGNAME(Default Table), COLUMNS(" _
        & LTrim(Str(Maxi)) & "), ROWS=(" & LTrim(Str(t.Rows.Count)) & ")" & ", HGUTTER(10000), VGUTTER(10583)"
    .TypeParagraph
    t.Cell(t.Rows.Count, t.Columns.Count).Select
    .Collapse wdCollapseEnd
    .TypeText "@Z_TBL_ROW_END = ": .TypeParagraph
    .TypeText "@Z_TBL_END = "
    .TypeParagraph 'Добавлено ПА 24.07.2007 на случай, если заголовки идут сразу после таблицы (без абзаца)
    t.ConvertToText wdSeparateByParagraphs
   
   End With
Next t

End Sub
Sub SplitHorCells(R As Row, Maxi As Long) 'не удалять - необходима для VenturaPrepTableXP
i = 1
ii = 1
Dim C As Cell
R.Cells(1).Select
For counter = 1 To R.Cells.Count + 1
    Set C = Selection.Cells(1)
    Selection.SelectColumn
    Set selColumn = Selection.Range
    lngClmn = selColumn.Cells(1).ColumnIndex
    If i = lngClmn Then
        ii = i
    Else
        i = IIf(lngClmn <> 1, lngClmn, Maxi + 1)
        C.Select
        With Selection
            .MoveLeft wdCell
            .Collapse
            .SelectCell
            .MoveEnd wdCharacter, -1
            .Cut
            .Cells.Split 1, i - ii
            .Collapse wdCollapseStart
            
            .Paste
            For i = i - ii To 2 Step -1
                .MoveRight wdCell
                .TypeText "@Z_TBL_CELL_BEG = HJOINED"
                .TypeParagraph
                .TypeText "@Z_TBL_CELL_END = "
            Next i
            'i = C.ColumnIndex
            SplitHorCells R, Maxi
            Exit Sub
        End With
    End If
    If i < Maxi Then
        i = i + 1
    Else
        i = 1
    End If
    C.Select: Selection.MoveRight wdCell
Next counter

End Sub

Sub HighlightEmptuCells()
Dim C As Cell
For Each C In Selection.Cells
    If C.Range.Text = Chr(13) & Chr(7) Then
    C.Range.Rows(1).Range.HighlightColorIndex = wdRed
    End If
Next C
End Sub
'    For Each R In T.Rows
'        R.Cells(1).Select: Selection.Collapse wdCollapseStart
'        If Not R.Index = 2 Then Selection.TypeText "@Z_TBL_ROW_END = ": Selection.TypeParagraph
'        Selection.TypeText "@Z_TBL_ROW_BEG = ": Selection.TypeParagraph
'        If R.Cells.Count < MaxI Then
'            R.Select: Selection.MoveDown wdLine, 1, wdExtend
'            Set selRows = Selection.Range
'            For Each C In selRows.Cells
'                C.Select: Selection.SelectColumn
'                Set selColumn = Selection.Range
'                lngClmn = selColumn.Cells(1).ColumnIndex
'                If i = lngClmn Then
'                    ii = i
'                Else
'                    i = IIf(lngClmn <> 1, lngClmn, MaxI + 1)
'                    C.Select
'                    With Selection
'                    .MoveLeft wdCell
'                    .Cut
'                    .Cells.Split 1, i - ii
'                    .Collapse wdCollapseStart
'
'                    .Paste
'                    For i = i - ii To 2 Step -1
'                        .MoveRight wdCell
'                        .TypeText "@Z_TBL_CELL_BEG = HJOINED"
'                        .TypeParagraph
'                        .TypeText "@Z_TBL_CELL_END = "
'                    Next i
'                    i = C.ColumnIndex
'                    a = R.Index: Set R = T.Rows(a - 1)
'                    Exit For
'                    End With
'                End If
'                If i < MaxI Then
'                    i = i + 1
'                Else
'                    i = 1
'                End If
'            Next C
'        End If
'    Next R

Sub GoToBookmark()
Attribute GoToBookmark.VB_Description = "Macro recorded 02.12.2006 by Artem"
Attribute GoToBookmark.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro2"
'
' Macro2 Macro
' Macro recorded 02.12.2006 by Artem
' Необходимо перед каждым переходом на скрытую закладку переключаться на режим редактирования сносок
' если закладка в тексте - Ворд её находит и из этого режима, но потом нужно опять переключаться
' В сносках маркер должен ставиться перед текстом сноски, иначе Вентура падает
    
ActiveWindow.View.SplitSpecial = wdPaneNone
ActiveWindow.ActivePane.View.Type = wdNormalView
ActiveWindow.View.SplitSpecial = wdPaneFootnotes
        
Selection.GoTo What:=wdGoToBookmark, Name:="_Ref152761726"
End Sub

Sub FootnotesConvert()
' Сначала запускать FieldsConvert
Dim F As Footnote
Dim FR As Range
If ActiveDocument.Footnotes.Count < 1 Then Exit Sub
ActiveWindow.View.SplitSpecial = wdPaneNone
ActiveWindow.ActivePane.View.Type = wdNormalView
Selection.HomeKey Unit:=wdStory, Extend:=wdMove
ActiveWindow.View.SplitSpecial = wdPaneFootnotes

For Each F In ActiveDocument.Footnotes
    ActiveWindow.Panes(2).Activate
    With Selection
        .EscapeKey
        .MoveRight wdWord
        Set FR = Selection.Range
        FR.Start = .Start
        If Not F.Index = ActiveDocument.Footnotes.Count Then
            .GoTo What:=wdGoToFootnote, Which:=wdGoToAbsolute, Count:=F.Index + 1 ' доделать
            .MoveLeft wdWord
        Else
            .EndKey Unit:=wdStory
        End If
        FR.End = .End
        FR.Select
        With .Find  'Замена параграфов на разрывы строк, если кто-нибудь воткнет в сноску стихи
            .ClearFormatting
            .Text = "^p"
            .Replacement.ClearFormatting
            .Replacement.Text = "<R>"
            .Forward = True: .Wrap = wdFindStop
            .Execute Replace:=wdReplaceAll
            .Text = "^l"
            .Replacement.Text = "<R>"
            .Forward = True: .Wrap = wdFindStop
            .Execute Replace:=wdReplaceAll
        End With
        FR.Select
        If Not InStr(.Text, Chr(13)) = 0 Then
            .MoveLeft wdCharacter, 1, wdExtend
        End If
        t = .Text '.Copy Изменил ПА 19.12.07 так как PasteAndFormat не работает
    End With
    ActiveWindow.Panes(1).Activate
    With Selection
        .GoTo What:=wdGoToFootnote, Which:=wdGoToNext
'Изменил ПА 19.12.07 так как PasteAndFormat не работает
'        .Font.Reset: .TypeText "<$F": .PasteAndFormat (wdFormatPlainText): .TypeText ">" 'здесь можно сразу втавлять стиль сноски?
        .Font.Reset: .TypeText "<$F" & t & ">"
    End With
    ActiveWindow.Panes(2).Activate
    Selection.GoTo What:=wdGoToFootnote, Which:=wdGoToNext
Next F
For Each F In ActiveDocument.Footnotes
    F.Delete
Next F
End Sub
Sub FieldsConvert()
Attribute FieldsConvert.VB_Description = "Macro recorded 02.12.2006 by Artem"
Attribute FieldsConvert.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro1 Macro
' Macro recorded 02.12.2006 by Artem
' Необходимо добавить проверку на >> - получается когда тег поля без точки оказывается в конце сноски
Dim Ref As String, F As Field, tag As Range
Selection.HomeKey Unit:=wdStory, Extend:=wdMove
ActiveWindow.View.SplitSpecial = wdPaneNone
ActiveWindow.ActivePane.View.Type = wdNormalView
If ActiveDocument.Footnotes.Count > 0 Then ActiveWindow.View.SplitSpecial = wdPaneFootnotes
ActiveWindow.Panes(1).Activate
For Each F In ActiveDocument.Fields
    If F.Type = wdFieldPageRef Then
        F.Update
        F.ShowCodes = True
        F.Select: Selection.Collapse wdCollapseStart
        Selection.MoveRight Unit:=wdWord, Count:=2
        Selection.MoveRight Unit:=wdWord, Count:=1, Extend:=wdExtend
        Ref = LTrim(RTrim(Selection.Text))
        F.Select: Selection.Delete
        With Selection
            .Collapse wdCollapseStart
            .Font.Reset: .TypeText "<$R[P#,": .TypeText Ref: .TypeText "]>"
        End With
        If ActiveDocument.Footnotes.Count > 0 Then ActiveWindow.Panes(2).Activate
        Selection.HomeKey Unit:=wdStory, Extend:=wdMove
        Selection.GoTo What:=wdGoToBookmark, Name:=Ref
        Selection.HomeKey Unit:=wdLine
        Selection.MoveRight Unit:=wdWord, Count:=1
        With Selection
            .MoveRight wdCharacter, 4, wdExtend
            If Not .Text = "<$M[" Then
                .Collapse wdCollapseStart
                .Font.Reset: .TypeText "<$M[": .TypeText Ref: .TypeText "]>"
            End If
        End With
        ActiveWindow.Panes(1).Activate
    End If
Next F
If ActiveDocument.Footnotes.Count > 0 Then
    ActiveWindow.Panes(2).Activate: ActiveWindow.Panes(2).View.ShowFieldCodes = True
    Selection.HomeKey Unit:=wdStory
    Selection.GoTo What:=wdGoToField, Which:=wdGoToNext, Count:=1, Name:= _
        "pageref"
    Do While Not Selection.Range.Start = 0
        Selection.MoveRight Unit:=wdWord, Count:=1, Extend:=wdExtend: Set tag = Selection.Range
        Selection.Collapse wdCollapseStart:  Selection.MoveRight Unit:=wdWord, Count:=2
        Selection.MoveRight Unit:=wdWord, Count:=1, Extend:=wdExtend
        Ref = LTrim(RTrim(Selection.Text))
        tag.Delete
        With Selection
            .Collapse wdCollapseStart
            .Font.Reset: .TypeText "<$R[P#,": .TypeText Ref: .TypeText "]>"
        End With
        Selection.HomeKey Unit:=wdStory, Extend:=wdMove
        Selection.GoTo What:=wdGoToBookmark, Name:=Ref
        'Selection.HomeKey Unit:=wdLine
        'Selection.MoveRight Unit:=wdWord, Count:=1
        With Selection
            .MoveRight wdCharacter, 4, wdExtend
            If Not .Text = "<$M[" Then
                .Collapse wdCollapseStart
                .Font.Reset: .TypeText "<$M[": .TypeText Ref: .TypeText "]>"
            End If
        End With
        ActiveWindow.Panes(2).Activate
        Selection.HomeKey Unit:=wdStory
        Selection.GoTo What:=wdGoToField, Which:=wdGoToNext, Count:=1, Name:= _
        "pageref"
    Loop
End If
End Sub
Sub Tabs_in_List()
Attribute Tabs_in_List.VB_Description = "Macro recorded 09.01.2007 by Artem"
Attribute Tabs_in_List.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Tabs_in_List"
'
' Tabs_in_List Macro
' Macro recorded 09.01.2007 by Artem
' Меняет пробелы на табуляторы внутри любого списка (обрабатывает только выделенный фрагмент)

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(^013[0-9]@.)( )"
        .Replacement.Text = "\1^t"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    
End Sub
Sub postal_list_format()
Attribute postal_list_format.VB_Description = "Macro recorded 10.04.2007 by Artem\r\nДобавляет Глав. Бухгалтера и рассыльщика на каждой 4 стр."
Attribute postal_list_format.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.postal_list_format"
'
' postal_list_format Macro
' Macro recorded 10.04.2007 by Artem
' Добавляет Глав. Бухгалтера и рассыльщика на каждой 4 стр.
'
    Application.Browser.Next
    Application.Browser.Next
    Application.Browser.Next
    Application.Browser.Next
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.InsertRowsAbove 1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.Paste
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.InsertBreak Type:=wdPageBreak
End Sub

Sub SplitDocToTxt()
Dim i
Dim SourceFileName$
Dim TargetFileName$
Dim SourcePath$
SourcePath$ = ActiveDocument.Path
CurDoc$ = WordBasic.[WindowName$]()
SourceFileName$ = CurDoc$
i = 1
WordBasic.EndOfDocument: WordBasic.InsertBreak
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditFind Find:="^m", Direction:=0, MatchCase:=0, WholeWord:=0, PatternMatch:=0, SoundsLike:=0, Format:=0, Wrap:=0, FindAllWordForms:=0
While WordBasic.EditFindFound()
    TargetFileName$ = SourceFileName$ + "_" + WordBasic.[LTrim$](Str(i)) + ".txt"
    If i < 100 Then TargetFileName$ = SourceFileName$ + "_0" + WordBasic.[LTrim$](Str(i)) + ".txt"
    If i < 10 Then TargetFileName$ = SourceFileName$ + "_00" + WordBasic.[LTrim$](Str(i)) + ".txt"
    WordBasic.EditClear: WordBasic.EditClear
    WordBasic.StartOfDocument 1
    WordBasic.EditCut
    WordBasic.FileNewDefault
    Selection.TypeText "@Z_STYLE70 = ": Selection.TypeParagraph
    WordBasic.EditPaste
    ActiveDocument.SaveAs FileName:=SourcePath$ & "\" & TargetFileName$, FileFormat:=wdFormatText, _
        LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword _
        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
        False, Encoding:=1251, InsertLineBreaks:=False, AllowSubstitutions:=False _
        , LineEnding:=wdCRLF
    'WordBasic.FileSaveAs Name:=TargetFileName$, Format:=2, LockAnnot:=0, Password:="", AddToMru:=1, WritePassword:="", RecommendReadOnly:=0, EmbedFonts:=0, NativePictureFormat:=0, FormsData:=0, SaveAsAOCELetter:=0
    WordBasic.FileClose 2
    i = i + 1
    WordBasic.Activate CurDoc$
    WordBasic.RepeatFind
Wend
End Sub
Sub TablesExtract()
Dim count_
Dim i
SourcePath$ = ActiveDocument.Path
CurDoc$ = WordBasic.[WindowName$]()
WordBasic.FileNew Template:="Normal", NewTemplate:=0
NewDoc$ = WordBasic.[WindowName$]()

WordBasic.Activate CurDoc$
WordBasic.StartOfDocument
WordBasic.EditBookmark "Temp", Add:=1
count_ = 0
If WordBasic.SelInfo(12) = -1 Then count_ = 1
WordBasic.WW7_EditGoTo Destination:="t"
While WordBasic.CmpBookmarks("\Sel", "Temp") <> 0
    WordBasic.EditBookmark "Temp", Add:=1
    count_ = count_ + 1
    WordBasic.RepeatFind
Wend
WordBasic.EditBookmark "Temp", Delete:=1
WordBasic.StartOfDocument
WordBasic.WW7_EditGoTo Destination:="t"
For i = 1 To count_
    WordBasic.TableSelectTable
    WordBasic.EditCut
    WordBasic.Activate NewDoc$
    WordBasic.EditPaste
    WordBasic.InsertBreak
    WordBasic.Activate CurDoc$
    WordBasic.RepeatFind
Next i
WordBasic.Activate NewDoc$
NewDocName$ = CurDoc$ & "_tbls.doc"
ActiveDocument.SaveAs FileName:=SourcePath$ & "\" & NewDocName$, FileFormat:= _
        wdFormatDocument, LockComments:=False, Password:="", AddToRecentFiles:= _
        True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:= _
        False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False
WordBasic.Activate CurDoc$
OldDocName$ = CurDoc$ & "_noTbls.doc"
ActiveDocument.SaveAs FileName:=SourcePath$ & "\" & OldDocName$, FileFormat:= _
        wdFormatDocument, LockComments:=False, Password:="", AddToRecentFiles:= _
        True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:= _
        False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False
ActiveDocument.Close wdSaveChanges
WordBasic.MsgBox "There are" + Str(count_) + " tables in the document"
CurDoc$ = NewDocName$
WordBasic.Activate CurDoc$
End Sub

Sub Markup07()
Dim oldscp
Dim i
Dim sn$
Dim sd$
Dim p
Dim a$
Dim p1
Dim symb$
Dim Inp$


'================================================================
'Создан 07.06.2007 на основе Markup06 ПА:
'правильно понимает стили шрифта
'быстрее делает вставку фигурных скобок в индексы
'не вставляет тэги п/ж и курсива на пробелы
'При обработке Symbol вставляет коды вместо символов - для совместимости с новым вордом
'На будущее - Инструкция Find с включенными Wildcards в этом макросе не работает!!!
'================================================================

'================================================================
'Создан 25.04.2007 на основе Markup04 ПА: добавлена совместимость с IndexEntryforWordinsert06 (индексными полями, отмеченными п/ж)
'================================================================

'==============================================
' Инициализация вентурных тэгов - добавлено 19.03.09 ПА для возможности вставлять Вентурные тэги до прохождения Markup
'==============================================
tagOpen = "[VENT:T:O]"
tagClose = "[VENT:T:C]"


' Запоминает значение Smart cut and paste и утанавливает его на 0
Dim toe As Object: Set toe = WordBasic.DialogRecord.ToolsOptionsEdit(False)
WordBasic.CurValues.ToolsOptionsEdit toe
oldscp = toe.SmartCutPaste
' MsgBox Str$(oldscp)
WordBasic.ToolsOptionsEdit SmartCutPaste:=0



WordBasic.EndOfDocument
WordBasic.ResetChar ' исправлено 25.04.2007 ПА: чтобы не зацикливался на последнем абзаце, если тот выделен п/ж или курсивом (не работает)
WordBasic.InsertPara
WordBasic.Style "Normal"
WordBasic.InsertPara
WordBasic.Insert "$"


WordBasic.ShowAll 0

' Сначала убьем все автоматические списки, чтобы буллиты и цифры не перелезли в верстку
' ========================================================================
For Each List In ActiveDocument.Lists
    List.RemoveNumbers
Next


' Знак больше и меньше
'============================================================
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditReplace Find:="<", Replace:="<<", Direction:=0, ReplaceAll:=1, Format:=0, Wrap:=0, PatternMatch:=0
WordBasic.EditReplace Find:=">", Replace:=">>", Direction:=0, ReplaceAll:=1, Format:=0, Wrap:=0

'Так лучше, но почему-то не всегда срабатывает (номера индексов) ПА
'WordBasic.StartOfDocument
'WordBasic.ShowAll 1
'WordBasic.EditFindClearFormatting
'WordBasic.EditReplace Find:="(\<)([0-9]@)(\>)", Replace:="{\2}", Direction:=0, MatchCase:=0, WholeWord:=0, PatternMatch:=1, SoundsLike:=0, ReplaceAll:=1, Format:=0, Wrap:=0, FindAllWordForms:=0
'
'WordBasic.ShowAll 0

' Элементы индекса
WordBasic.StartOfDocument
WordBasic.ShowAll 1
WordBasic.EditFindClearFormatting
WordBasic.EditFind Find:="<^#", Direction:=0, Format:=0, Wrap:=0
While WordBasic.EditFindFound()
    WordBasic.CharLeft
    WordBasic.EditClear
    WordBasic.Insert Chr(123)
    WordBasic.RepeatFind
Wend

WordBasic.StartOfDocument
WordBasic.EditFind Find:="^#>", Direction:=0, Format:=0, Wrap:=0
While WordBasic.EditFindFound()
    WordBasic.CharRight
    WordBasic.CharLeft
    WordBasic.EditClear
    WordBasic.Insert Chr(125)
    WordBasic.RepeatFind
Wend

WordBasic.ShowAll 0

'==============================================
' Восстановление вентурных тэгов - добавлено 19.03.09 ПА для возможности вставлять Вентурные тэги до прохождения Markup
'==============================================
Selection.HomeKey Unit:=wdStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
    .Text = "[VENT:T:O]"
    .Replacement.Text = "<"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll

Selection.HomeKey Unit:=wdStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
    .Text = "[VENT:T:C]"
    .Replacement.Text = ">"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll

' Курсив.
' Хе. Оказывается, smart cut and paste должен быть отключен!
' Дописал
'===================================================================

WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditFindFont Italic:=1
WordBasic.EditFind Find:="", Direction:=0, Format:=1, Wrap:=0, PatternMatch:=0
While WordBasic.EditFindFound()

'Проверка на захват знака абзаца (заодно избавляет от глюков
'кода знак абзаца отформатирован иначе чем строка перед ним)
If Asc(WordBasic.[Right$](WordBasic.[Selection$](), 1)) = 13 Then
    If Len(WordBasic.[Selection$]()) = 1 Then GoTo repFindItalic
    WordBasic.CharLeft 1, 1
End If
    If Len(WordBasic.[Selection$]()) = 1 And WordBasic.[Selection$]() = Chr(32) Then GoTo repFindItalic ' дописано 08.06.07 ПА: что бы не вставлял лишних тэгов на пробелы
    WordBasic.EditCut
    WordBasic.ResetChar
    WordBasic.Insert "<I>"
    WordBasic.EditPaste
    WordBasic.ResetChar
    WordBasic.Insert "<I*>"
repFindItalic:
    WordBasic.ResetChar
    WordBasic.RepeatFind
Wend

' Полужирный
'===========================================================
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditFindFont Bold:=1
WordBasic.EditFind Find:="", Direction:=0, Format:=1, Wrap:=0
While WordBasic.EditFindFound()

'Проверка на захват знака абзаца (заодно избавляет от глюков
'кода знак абзаца отформатирован иначе чем строка перед ним)
If Asc(WordBasic.[Right$](WordBasic.[Selection$](), 1)) = 13 Then
    If Len(WordBasic.[Selection$]()) = 1 Then GoTo repFindBold
    WordBasic.CharLeft 1, 1
End If
    If Len(WordBasic.[Selection$]()) = 1 And WordBasic.[Selection$]() = Chr(32) Then GoTo repFindBold ' дописано 08.06.07 ПА: что бы не вставлял лишних тэгов на пробелы
    WordBasic.EditCut
    WordBasic.ResetChar
    WordBasic.Insert "<B>"
    WordBasic.EditPaste
    WordBasic.ResetChar
    WordBasic.Insert "<W0>"
repFindBold:
    WordBasic.ResetChar
    WordBasic.RepeatFind
Wend

' Подчеркивание
'===========================================================
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditFindFont Underline:=1
WordBasic.EditFind Find:="", Direction:=0, Format:=1, Wrap:=0
While WordBasic.EditFindFound()
    WordBasic.EditCut
    WordBasic.ResetChar
    WordBasic.Insert "<U>"
    WordBasic.EditPaste
    WordBasic.ResetChar
    WordBasic.Insert "<U*>"
    WordBasic.ResetChar
    WordBasic.RepeatFind
Wend



' Символ
'===========================================================
' Изменено 09.06.2007 АП: вставляет коды вместо символов - для совместимости с новым вордом
'==================================

WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditFindFont Font:="Symbol"
WordBasic.EditFind Find:="", Direction:=0, Format:=1, Wrap:=0
While WordBasic.EditFindFound()
    If WordBasic.[Selection$]() <> Chr(13) Then
        If InStr(WordBasic.[Selection$](), Chr(13)) > 0 Then WordBasic.CharLeft 1, 1
        symb$ = WordBasic.[Selection$]()
        WordBasic.EditClear
        WordBasic.Insert "<F" + Chr(34) + "Symbol" + Chr(34) + ">"
        For i = 1 To Len(symb$)
            Inp$ = WordBasic.[LTrim$](Str(AscB(Mid(symb$, i, 1))))
            If AscB(Mid(symb$, i, 1)) < 100 Then Inp$ = "0" + Inp$
            WordBasic.Insert "<@" + Inp$ + ">"
        Next i
        WordBasic.Insert "<F255>"
        WordBasic.ResetChar
    Else
        WordBasic.Style "Default Paragraph Font"
    End If
    WordBasic.RepeatFind
Wend


' Стили
'===========================================================
'Расстановка стилей шрифта - добавлено 07.06.2007 ПА
For i = 1 To WordBasic.CountStyles(0, 0)
    sn$ = WordBasic.[StyleName$](i, 0, 0)
    sd$ = WordBasic.[StyleDesc$](sn$)
    sd$ = WordBasic.[Left$](sd$, InStr(sd$, "+"))
    If sd$ = "Default Paragraph Font +" Then 'Если стиль основан на стиле шрифта то есть сам он тоже стиль шрифта
        WordBasic.StartOfDocument
        WordBasic.EditFindClearFormatting
        WordBasic.EditFindStyle Style:=sn$
        WordBasic.EditFind Find:="", Direction:=0, Format:=1, Wrap:=0
        While WordBasic.EditFindFound()
            If WordBasic.[Selection$]() <> Chr(13) Then
                WordBasic.Style "Default Paragraph Font"
                WordBasic.Bold 0: WordBasic.Italic 0
                WordBasic.EditCut
                WordBasic.Style "Default Paragraph Font"
                WordBasic.Bold 0: WordBasic.Italic 0
                WordBasic.Insert "<$[" + sn$ + ">"
                WordBasic.EditPaste
                WordBasic.Style "Default Paragraph Font"
                WordBasic.Bold 0: WordBasic.Italic 0
                WordBasic.Insert "<$]" + sn$ + ">"
            End If
            WordBasic.RepeatFind
        Wend
    End If
Next i

'Расстановка стилей абзаца
WordBasic.StartOfDocument
p = 0
While p = 0
    If WordBasic.[Selection$]() = "$" Then
        p = 1
    Else
        If WordBasic.[Selection$]() <> "@" Then
' And StyleName$() <> "Normal" Then
            WordBasic.SelType 1
            WordBasic.Style "Default Paragraph Font"
            WordBasic.Insert "@" + WordBasic.[StyleName$]() + " = "
        End If
        WordBasic.ParaDown
    End If
Wend

WordBasic.StartOfDocument
WordBasic.Insert "@Z_STYLE70 = ": WordBasic.InsertPara


'Цифровой пробел
'============================================================
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditFind Find:="^#^s^#", Direction:=0, Format:=0, Wrap:=0, PatternMatch:=0
While WordBasic.EditFindFound()
    WordBasic.CharLeft 1
    WordBasic.CharRight 1
    WordBasic.CharRight 1, 1
    WordBasic.EditClear
    WordBasic.Insert "<|>"
    WordBasic.RepeatFind
Wend

' Длинное тире с пробелом
'============================================================
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditReplace Find:=" ^+", Replace:="<N><@151>", Direction:=0, ReplaceAll:=1, Format:=0, Wrap:=0


' Длинное и короткое тире
'============================================================
WordBasic.EditReplace Find:="^+", Replace:="<@151>", Direction:=0, ReplaceAll:=1, Format:=0, Wrap:=0
WordBasic.EditReplace Find:="^=", Replace:="<@150>", Direction:=0, ReplaceAll:=1, Format:=0, Wrap:=0

' Неразрывный пробел
'============================================================
WordBasic.EditReplace Find:="^s", Replace:="<N>", Direction:=0, ReplaceAll:=1, Format:=0, Wrap:=0

' Мягкие переносы
'===========================================================
WordBasic.EditFindClearFormatting
WordBasic.EditReplace Find:="^-", Replace:="<->", Direction:=0, ReplaceAll:=1, Format:=0, Wrap:=0

' Метки предметного указателя - изменено 25.04.2007 ПА
'===========================================================

WordBasic.EditFindClearFormatting
WordBasic.ShowAll 1
WordBasic.StartOfDocument
WordBasic.EditFind Find:="^019", Direction:=0, Format:=0, Wrap:=0
While WordBasic.EditFindFound()
    a$ = WordBasic.[Selection$]()
    p = InStr(a$, Chr(34))
    a$ = Mid(a$, p + 1)
    p = InStr(a$, Chr(34))
    p1 = InStr(a$, "\b") 'нахождение маркера п/ж
    a$ = WordBasic.[Left$](a$, p - 1)
    If p1 <> 0 And p1 > p Then 'Если есть маркер и он не часть текста
        WordBasic.Insert "<$I\b" + a$ + ">" 'добавляет маркер п/ж
    Else
        WordBasic.Insert "<$I" + a$ + ">"
    End If
    WordBasic.RepeatFind
Wend

' Суперскрипт
'===========================================================
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditFindFont Superscript:=1
WordBasic.EditFind Find:="", Direction:=0, Format:=1, Wrap:=0, MatchCase:=0
While WordBasic.EditFindFound()
    WordBasic.EditCut
    WordBasic.Insert "<^>"
    WordBasic.EditPaste
    WordBasic.Insert "<^*>"
    WordBasic.ResetChar
    WordBasic.RepeatFind
Wend

' Субскрипт
'===========================================================
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditFindFont Subscript:=1
WordBasic.EditFind Find:="", Direction:=0, Format:=1, Wrap:=0
While WordBasic.EditFindFound()
    WordBasic.EditCut
    WordBasic.Insert "<V>"
    WordBasic.EditPaste
    WordBasic.Insert "<^*>"
    WordBasic.ResetChar
    WordBasic.RepeatFind
Wend

' Разрывы строк
'==========================================================
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditReplace Find:="^l", Replace:="<R>", Direction:=0, ReplaceAll:=1, Format:=0, Wrap:=0

WordBasic.EndOfDocument
WordBasic.EditClear -2


' Восстанавливает старое значение Smart cut and paste
WordBasic.ToolsOptionsEdit SmartCutPaste:=oldscp

End Sub



Sub ProcessAllTables()
TablesExtract
WordBasic.Activate CurDoc$
VenturaPrepTableXP
WordBasic.Activate CurDoc$
Markup07
WordBasic.Activate CurDoc$
SplitDocToTxt

End Sub
Sub SelectCurrentLevel()
Selection.GoTo What:=wdGoToBookmark, Name:="\HeadingLevel"
End Sub

Sub Markup_pre_fields()
'Создан из Markup07 для работы с макросами FieldsConvert и FootnotesConvert
'Сначала прибивает все знаки больше-меньше
'Затем запускать макросы FieldsConvert и FootnotesConvert
'Затем запускать Markup_post_fields
'ПА 19.12.07
ActiveWindow.View.SplitSpecial = wdPaneNone
ActiveWindow.ActivePane.View.Type = wdNormalView
Selection.HomeKey Unit:=wdStory, Extend:=wdMove
ActiveWindow.View.SplitSpecial = wdPaneFootnotes

ActiveWindow.Panes(1).Activate

' Знак больше и меньше
'============================================================
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditReplace Find:="<", Replace:="<<", Direction:=0, ReplaceAll:=1, Format:=0, Wrap:=0, PatternMatch:=0
WordBasic.EditReplace Find:=">", Replace:=">>", Direction:=0, ReplaceAll:=1, Format:=0, Wrap:=0

ActiveWindow.Panes(2).Activate

' Знак больше и меньше
'============================================================
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditReplace Find:="<", Replace:="<<", Direction:=0, ReplaceAll:=1, Format:=0, Wrap:=0, PatternMatch:=0
WordBasic.EditReplace Find:=">", Replace:=">>", Direction:=0, ReplaceAll:=1, Format:=0, Wrap:=0

End Sub
Sub Markup_post_fields()
'Создан из Markup07 для работы с макросами FieldsConvert и FootnotesConvert
'Сначала прибивает все знаки больше-меньше
'Затем запускать макросы FieldsConvert и FootnotesConvert
'Затем запускать Markup_post_fields
'ПА 19.12.07

Dim oldscp
Dim i
Dim sn$
Dim sd$
Dim p
Dim a$
Dim p1
Dim symb$
Dim Inp$

'================================================================
'Создан 07.06.2007 на основе Markup06 ПА:
'правильно понимает стили шрифта
'быстрее делает вставку фигурных скобок в индексы
'не вставляет тэги п/ж и курсива на пробелы
'При обработке Symbol вставляет коды вместо символов - для совместимости с новым вордом
'================================================================


'================================================================
'Создан 25.04.2007 на основе Markup04 ПА: добавлена совместимость с IndexEntryforWordinsert06 (индексными полями, отмеченными п/ж)
'================================================================

' Запоминает значение Smart cut and paste и утанавливает его на 0
Dim toe As Object: Set toe = WordBasic.DialogRecord.ToolsOptionsEdit(False)
WordBasic.CurValues.ToolsOptionsEdit toe
oldscp = toe.SmartCutPaste
' MsgBox Str$(oldscp)
WordBasic.ToolsOptionsEdit SmartCutPaste:=0
'----------------------------------------------------------------
If 1 = 2 Then
For i = 1 To WordBasic.CountStyles(0, 0)
    sn$ = WordBasic.[StyleName$](i, 0, 0)
    If sn$ <> "Default Paragraph Font" Then
        WordBasic.FormatStyle Name:=sn$, Define:=1
        WordBasic.FormatDefineStyleFont Bold:=0, Italic:=0
'       FormatDefineStyleNumbers .Remove
    End If
Next i
End If

WordBasic.EndOfDocument
WordBasic.ResetChar ' исправлено 25.04.2007 ПА: чтобы не зацикливался на последнем абзаце, если тот выделен п/ж или курсивом (не работает)
WordBasic.InsertPara
WordBasic.Style "Normal"
WordBasic.InsertPara
WordBasic.Insert "$"


WordBasic.ShowAll 0


WordBasic.StartOfDocument
WordBasic.ShowAll 1
WordBasic.EditFindClearFormatting
WordBasic.EditFind Find:="<^#", Direction:=0, Format:=0, Wrap:=0
While WordBasic.EditFindFound()
    WordBasic.CharLeft
    WordBasic.EditClear
    WordBasic.Insert Chr(123)
    WordBasic.RepeatFind
Wend

WordBasic.StartOfDocument
WordBasic.EditFind Find:="^#>", Direction:=0, Format:=0, Wrap:=0
While WordBasic.EditFindFound()
    WordBasic.CharRight
    WordBasic.CharLeft
    WordBasic.EditClear
    WordBasic.Insert Chr(125)
    WordBasic.RepeatFind
Wend

WordBasic.ShowAll 0

' Курсив.
' Хе. Оказывается, smart cut and paste должен быть отключен!
' Дописал
'===================================================================

WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditFindFont Italic:=1
WordBasic.EditFind Find:="", Direction:=0, Format:=1, Wrap:=0, PatternMatch:=0
While WordBasic.EditFindFound()

'Проверка на захват знака абзаца (заодно избавляет от глюков
'кода знак абзаца отформатирован иначе чем строка перед ним)
If Asc(WordBasic.[Right$](WordBasic.[Selection$](), 1)) = 13 Then
    If Len(WordBasic.[Selection$]()) = 1 Then GoTo repFindItalic
    WordBasic.CharLeft 1, 1
End If
    If Len(WordBasic.[Selection$]()) = 1 And WordBasic.[Selection$]() = Chr(32) Then GoTo repFindItalic ' дописано 08.06.07 ПА: что бы не вставлял лишних тэгов на пробелы
    WordBasic.EditCut
    WordBasic.ResetChar
    WordBasic.Insert "<I>"
    WordBasic.EditPaste
    WordBasic.ResetChar
    WordBasic.Insert "<I*>"
repFindItalic:
    WordBasic.ResetChar
    WordBasic.RepeatFind
Wend

' Полужирный
'===========================================================
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditFindFont Bold:=1
WordBasic.EditFind Find:="", Direction:=0, Format:=1, Wrap:=0
While WordBasic.EditFindFound()

'Проверка на захват знака абзаца (заодно избавляет от глюков
'кода знак абзаца отформатирован иначе чем строка перед ним)
If Asc(WordBasic.[Right$](WordBasic.[Selection$](), 1)) = 13 Then
    If Len(WordBasic.[Selection$]()) = 1 Then GoTo repFindBold
    WordBasic.CharLeft 1, 1
End If
    If Len(WordBasic.[Selection$]()) = 1 And WordBasic.[Selection$]() = Chr(32) Then GoTo repFindBold ' дописано 08.06.07 ПА: что бы не вставлял лишних тэгов на пробелы
    WordBasic.EditCut
    WordBasic.ResetChar
    WordBasic.Insert "<B>"
    WordBasic.EditPaste
    WordBasic.ResetChar
    WordBasic.Insert "<W0>"
repFindBold:
    WordBasic.ResetChar
    WordBasic.RepeatFind
Wend

' Подчеркивание
'===========================================================
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditFindFont Underline:=1
WordBasic.EditFind Find:="", Direction:=0, Format:=1, Wrap:=0
While WordBasic.EditFindFound()
    WordBasic.EditCut
    WordBasic.ResetChar
    WordBasic.Insert "<U>"
    WordBasic.EditPaste
    WordBasic.ResetChar
    WordBasic.Insert "<U*>"
    WordBasic.ResetChar
    WordBasic.RepeatFind
Wend



' Символ
'===========================================================
' Изменено 09.06.2007 АП: вставляет коды вместо символов - для совместимости с новым вордом
'===========================================================
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditFindFont Font:="Symbol"
WordBasic.EditFind Find:="", Direction:=0, Format:=1, Wrap:=0
While WordBasic.EditFindFound()
    If WordBasic.[Selection$]() <> Chr(13) Then
        If InStr(WordBasic.[Selection$](), Chr(13)) > 0 Then WordBasic.CharLeft 1, 1
        symb$ = WordBasic.[Selection$]()
        WordBasic.EditClear
        WordBasic.Insert "<F" + Chr(34) + "Symbol" + Chr(34) + ">"
        For i = 1 To Len(symb$)
            Inp$ = WordBasic.[LTrim$](Str(AscB(Mid(symb$, i, 1))))
            If AscB(Mid(symb$, i, 1)) < 100 Then Inp$ = "0" + Inp$
            WordBasic.Insert "<@" + Inp$ + ">"
        Next i
        WordBasic.Insert "<F255>"
        WordBasic.ResetChar
    Else
        WordBasic.Style "Default Paragraph Font"
    End If
    WordBasic.RepeatFind
Wend

' Стили
'===========================================================
'Расстановка стилей шрифта - добавлено 07.06.2007 ПА
For i = 1 To WordBasic.CountStyles(0, 0)
    sn$ = WordBasic.[StyleName$](i, 0, 0)
    sd$ = WordBasic.[StyleDesc$](sn$)
    sd$ = WordBasic.[Left$](sd$, InStr(sd$, "+"))
    If sd$ = "Default Paragraph Font +" Then 'Если стиль основан на стиле шрифта то есть сам он тоже стиль шрифта
        WordBasic.StartOfDocument
        WordBasic.EditFindClearFormatting
        WordBasic.EditFindStyle Style:=sn$
        WordBasic.EditFind Find:="", Direction:=0, Format:=1, Wrap:=0
        While WordBasic.EditFindFound()
            If WordBasic.[Selection$]() <> Chr(13) Then
                WordBasic.Style "Default Paragraph Font"
                WordBasic.Bold 0: WordBasic.Italic 0
                WordBasic.EditCut
                WordBasic.Style "Default Paragraph Font"
                WordBasic.Bold 0: WordBasic.Italic 0
                WordBasic.Insert "<$[" + sn$ + ">"
                WordBasic.EditPaste
                WordBasic.Style "Default Paragraph Font"
                WordBasic.Bold 0: WordBasic.Italic 0
                WordBasic.Insert "<$]" + sn$ + ">"
            End If
            WordBasic.RepeatFind
        Wend
    End If
Next i

'Расстановка стилей абзаца
WordBasic.StartOfDocument
p = 0
While p = 0
    If WordBasic.[Selection$]() = "$" Then
        p = 1
    Else
        If WordBasic.[Selection$]() <> "@" Then
' And StyleName$() <> "Normal" Then
            WordBasic.SelType 1
            WordBasic.Style "Default Paragraph Font"
            WordBasic.Insert "@" + WordBasic.[StyleName$]() + " = "
        End If
        WordBasic.ParaDown
    End If
Wend

WordBasic.StartOfDocument
WordBasic.Insert "@Z_STYLE70 = ": WordBasic.InsertPara


'Цифровой пробел
'============================================================
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditFind Find:="^#^s^#", Direction:=0, Format:=0, Wrap:=0, PatternMatch:=0
While WordBasic.EditFindFound()
    WordBasic.CharLeft 1
    WordBasic.CharRight 1
    WordBasic.CharRight 1, 1
    WordBasic.EditClear
    WordBasic.Insert "<|>"
    WordBasic.RepeatFind
Wend

' Длинное тире с пробелом
'============================================================
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditReplace Find:=" ^+", Replace:="<N><@151>", Direction:=0, ReplaceAll:=1, Format:=0, Wrap:=0


' Длинное и короткое тире
'============================================================
WordBasic.EditReplace Find:="^+", Replace:="<@151>", Direction:=0, ReplaceAll:=1, Format:=0, Wrap:=0
WordBasic.EditReplace Find:="^=", Replace:="<@150>", Direction:=0, ReplaceAll:=1, Format:=0, Wrap:=0

' Неразрывный пробел
'============================================================
WordBasic.EditReplace Find:="^s", Replace:="<N>", Direction:=0, ReplaceAll:=1, Format:=0, Wrap:=0

' Мягкие переносы
'===========================================================
WordBasic.EditFindClearFormatting
WordBasic.EditReplace Find:="^-", Replace:="<->", Direction:=0, ReplaceAll:=1, Format:=0, Wrap:=0

' Метки предметного указателя - изменено 25.04.2007 ПА
'===========================================================

WordBasic.EditFindClearFormatting
WordBasic.ShowAll 1
WordBasic.StartOfDocument
WordBasic.EditFind Find:="^019", Direction:=0, Format:=0, Wrap:=0
While WordBasic.EditFindFound()
    a$ = WordBasic.[Selection$]()
    p = InStr(a$, Chr(34))
    a$ = Mid(a$, p + 1)
    p = InStr(a$, Chr(34))
    p1 = InStr(a$, "\b") 'нахождение маркера п/ж
    a$ = WordBasic.[Left$](a$, p - 1)
    If p1 <> 0 And p1 > p Then 'Если есть маркер и он не часть текста
        WordBasic.Insert "<$I\b" + a$ + ">" 'добавляет маркер п/ж
    Else
        WordBasic.Insert "<$I" + a$ + ">"
    End If
    WordBasic.RepeatFind
Wend

' Суперскрипт
'===========================================================
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditFindFont Superscript:=1
WordBasic.EditFind Find:="", Direction:=0, Format:=1, Wrap:=0, MatchCase:=0
While WordBasic.EditFindFound()
    WordBasic.EditCut
    WordBasic.Insert "<^>"
    WordBasic.EditPaste
    WordBasic.Insert "<^*>"
    WordBasic.ResetChar
    WordBasic.RepeatFind
Wend

' Субскрипт
'===========================================================
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditFindFont Subscript:=1
WordBasic.EditFind Find:="", Direction:=0, Format:=1, Wrap:=0
While WordBasic.EditFindFound()
    WordBasic.EditCut
    WordBasic.Insert "<V>"
    WordBasic.EditPaste
    WordBasic.Insert "<^*>"
    WordBasic.ResetChar
    WordBasic.RepeatFind
Wend

' Разрывы строк
'==========================================================
WordBasic.StartOfDocument
WordBasic.EditFindClearFormatting
WordBasic.EditReplace Find:="^l", Replace:="<R>", Direction:=0, ReplaceAll:=1, Format:=0, Wrap:=0

WordBasic.EndOfDocument
WordBasic.EditClear -2


' Восстанавливает старое значение Smart cut and paste
WordBasic.ToolsOptionsEdit SmartCutPaste:=oldscp

End Sub



Sub DiacriticsToCodes()
'
' DiacriticsToCodes Macro
' Macro recorded 17.03.2009 by Artem
' Перед запуском должны быть открыты 2 файла: текст и таблица подстановки
' Перед запуском должна быть выделена таблица подстановки
' Формат 1 колонка - оригинал
' 2 = Times New Roman Greek
' 3 = NewtonT
' 4 = NewtonE
' 5 = Times New Roman Baltic
' 6 = Код символа
' Символ должен стоять в колонке с тем шрифтом которому он соответствует
'
'==============================================
' Инициализация вентурных тэгов
'==============================================
tagOpen = "[VENT:T:O]"
tagClose = "[VENT:T:C]"

'
'==============================================
' Проверка: сколько окон открыто
'==============================================

If Not Windows.Count = 2 Then
    MsgBox "Перед запуском должны быть открыты 2 файла: текст и таблица подстановки"
    Exit Sub
End If
Dim docTbl As Document
Dim docText As Document
Set docTbl = Selection.Document
'winTblNum = docTbl.Windows(1).Index
'==============================================
' Проверка: выбрана ли таблица
'==============================================
If Selection.Tables.Count < 1 Then
    MsgBox "Перед запуском должна быть выделена таблица подстановки"
    Exit Sub
End If



If docTbl.Windows(1).Index = 1 Then
    Windows(2).Activate
Else
    Windows(1).Activate
End If
Set docText = Selection.Document

docTbl.Activate
'==============================================
' Очищаем таблицу от лишних пробелов и абзацев
'==============================================
    Selection.Tables(1).Select
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
  '=========================================================
  ' Превращаем таблицу в список с ; по их кол-ву будет определяться шрифт
  '=========================================================
    
    Selection.Rows.ConvertToText Separator:=wdSeparateByCommas, NestedTables:= _
        True
  '===============================================
  ' Убираем появившиеся пробелы
  '===============================================
        
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
  '===============================================
  ' Убираем лишние абзацы
  '===============================================
        
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    
    'Переключаемся во второе окно
    
'    docText.Activate
'
'
'  '=========================================================
'  ' Пересохраняем файл - ну его нах
'  '=========================================================
'
'SourcePath = ActiveDocument.Path
'NewDocName = docText.Name & "_process-diacritics"
'    ActiveDocument.SaveAs FileName:=SourcePath & "\" & NewDocName, _
'        FileFormat:=wdFormatDocument, LockComments:=False, Password:="", _
'        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
'        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
'        :=False, SaveAsAOCELetter:=False
'Set docText = Selection.Document

'docTbl.Activate

'================================================
'Начинаем работу
'================================================

Selection.EndKey Unit:=wdStory
Selection.TypeParagraph
Selection.TypeText "[EndOfFile]"

Selection.HomeKey Unit:=wdStory 'может и не надо?
Selection.EndKey Unit:=wdLine, Extend:=wdExtend
Selection.MoveLeft Unit:=wdCharacter, Extend:=wdExtend
Do While Not Selection = "[EndOfFile]" And Not Selection = Chr(13)
    
    'выбираем символ
    Selection.MoveLeft wdCharacter
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    substChr = AscW(Selection)
    
    'выбираем все ;
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdWord, Count:=1, Extend:=wdExtend
    i = Len(Selection)
    Select Case Len(Selection)
    Case 1
        strFontName = "Times New Roman Greek"
    Case 2
        strFontName = "NewtonT"
    
    Case 3
        strFontName = "NewtonE"

    Case 4
        strFontName = "Times New Roman Baltic"

    Case Else
    MsgBox "Где-то пропущен символ"
    End Select
    
    ' Забираем код символа
    Selection.EndKey Unit:=wdLine
    Selection.MoveLeft Unit:=wdWord, Count:=1, Extend:=wdExtend
    ChrCode = Selection
    
    'Переключаемся в текст
    docText.Activate
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(substChr)
        .Replacement.Text = tagOpen & "F""" & strFontName & """" & tagClose & Chr(ChrCode) & _
        tagOpen & "F255" & tagClose
        
  ' Этот вариант вставлял коды, но пришлось изменить из-за ошибки 255 (вентура грохает строчку, в которой есть @255)
  '      .Replacement.Text = tagOpen & "F""" & strFontName & """" & tagClose & tagOpen & "@" & ChrCode & tagClose & _
        tagOpen & "F255" & tagClose
        
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' Возвращаемся в таблицу
    docTbl.Activate
    Selection.MoveDown
    Selection.HomeKey
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.MoveLeft Unit:=wdCharacter, Extend:=wdExtend

Loop

Selection.Delete

End Sub


Sub viewCharCode()
'
' viewCharCode Macro
' Macro recorded 10.03.2009 by Artem
'

    Selection.HomeKey Unit:=wdLine
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    
    Inp$ = WordBasic.[LTrim$](Str(AscB(Selection.Text)))
            If Inp$ < 100 Then Inp$ = "0" + Inp$
            
    Selection.SelectRow
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeText Inp$
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveLeft Unit:=wdCell
    Selection.MoveLeft Unit:=wdCell
End Sub
Sub InputCharCode()
'
' viewCharCode Macro
' Macro recorded 10.03.2009 by Artem
'

    Selection.HomeKey Unit:=wdLine
    Selection.MoveRight Unit:=wdWord, Count:=1, Extend:=wdExtend
    
    Inp$ = Chr(CLng(Selection.Text))

            
    Selection.SelectRow
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeText Inp$
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveLeft Unit:=wdCell
End Sub

