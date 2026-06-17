Sub FlattenDocumentToText()
    ' Converts auto-numbered lists AND all fields (captions, cross-refs,
    ' TOC, index, page numbers, SEQ, etc.) into static text.
    Dim rngStory As Range
    
    ' 1. Lists ? text
    On Error Resume Next
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    On Error GoTo 0
    
    ' 2. Fields ? text, across every story (body, headers, footers, footnotes...)
    For Each rngStory In ActiveDocument.StoryRanges
        Do
            On Error Resume Next
            rngStory.Fields.Unlink
            On Error GoTo 0
            Set rngStory = rngStory.NextStoryRange
        Loop Until rngStory Is Nothing
    Next rngStory
    
    MsgBox "Done — lists and fields flattened to text.", vbInformation
End Sub
