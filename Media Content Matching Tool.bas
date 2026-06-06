Attribute VB_Name = "Module1"
' ===================== CLEAN TEXT FUNCTION =====================
Function CleanText(ByVal text As String) As String
    Dim c As Integer
    text = LCase(text)
    For c = 1 To 255
        If Not ((c >= 48 And c <= 57) Or (c >= 97 And c <= 122) Or c = 32) Then
            text = Replace(text, Chr(c), " ")
        End If
    Next c
    CleanText = Application.WorksheetFunction.Trim(text)
End Function

' ===================== STOPWORD REMOVAL =====================
Function RemoveStopWords(words As Variant) As Collection
    Dim stopWords As Variant
    stopWords = Array("the", "a", "an", "and", "or", "of", "in", "on", "at", "for", "with", "from", "by", "to")
    Dim result As New Collection, i As Long, w As String, sw As Variant, isStopWord As Boolean
    For i = LBound(words) To UBound(words)
        w = Trim(words(i))
        isStopWord = False
        For Each sw In stopWords
            If w = sw Then isStopWord = True: Exit For
        Next sw
        If Not isStopWord And Len(w) > 0 Then result.Add w
    Next i
    Set RemoveStopWords = result
End Function

' ===================== STEMMING FUNCTION =====================
Function StemWord(ByVal word As String) As String
    If Len(word) > 3 And Right(word, 1) = "s" Then
        StemWord = Left(word, Len(word) - 1)
    Else
        StemWord = word
    End If
End Function

' ===================== EXTRACT QUOTED TITLE =====================
Function ExtractQuotedTitle(ByVal text As String) As String
    Dim s As Long, e As Long
    If InStr(text, "Title:") > 0 And InStr(text, "and Series") > 0 Then
        s = InStr(text, "Title:") + 6
        e = InStr(text, "and Series")
        ExtractQuotedTitle = Trim(Mid(text, s, e - s))
    ElseIf InStr(text, """") > 0 Then
        s = InStr(text, """")
        e = InStr(s + 1, text, """")
        If e > s Then ExtractQuotedTitle = Mid(text, s + 1, e - s - 1): Exit Function
        ExtractQuotedTitle = text
    Else
        ExtractQuotedTitle = text
    End If
End Function

' ===================== EXTRACT PLATFORM/CHANNEL =====================
Function ExtractPlatformChannel(ByVal text As String) As String
    If InStr(text, "Platform:") > 0 Then
        ExtractPlatformChannel = Trim(Mid(text, InStr(text, "Platform:") + 9))
    ElseIf InStr(text, "Channel:") > 0 Then
        ExtractPlatformChannel = Trim(Mid(text, InStr(text, "Channel:") + 8))
    Else
        ExtractPlatformChannel = ""
    End If
End Function

' ===================== MAIN MACRO =====================
Sub MatchMovieTitles_SubstringFiltered()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim wsMovie As Worksheet, wsMaster As Worksheet, wsOutput As Worksheet
    Set wsMovie = ThisWorkbook.Sheets("Movie_Titles")
    Set wsMaster = ThisWorkbook.Sheets("Master_Titles")
    Set wsOutput = ThisWorkbook.Sheets("Match_Results")
    wsOutput.Cells.ClearContents
    wsOutput.Cells.FormatConditions.Delete

    ' Headers
    wsOutput.Range("A1:H1").Value = Array("Movie Title", "Status", "Best Match Title", "Match Score", "Asset Type", "Asset Code", "Platform/Channel", "List of Titles")

    Dim lastMovieRow As Long, lastMasterRow As Long
    lastMovieRow = wsMovie.Cells(wsMovie.Rows.count, 1).End(xlUp).Row
    lastMasterRow = wsMaster.Cells(wsMaster.Rows.count, 1).End(xlUp).Row

    Dim masterRawArr(), masterCleanArr(), masterWordArr(), assetTypeArr(), assetCodeArr()
    ReDim masterRawArr(2 To lastMasterRow)
    ReDim masterCleanArr(2 To lastMasterRow)
    ReDim masterWordArr(2 To lastMasterRow)
    ReDim assetTypeArr(2 To lastMasterRow)
    ReDim assetCodeArr(2 To lastMasterRow)

    Dim j As Long
    For j = 2 To lastMasterRow
        masterRawArr(j) = wsMaster.Cells(j, 1).Value
        assetTypeArr(j) = wsMaster.Cells(j, 2).Value
        assetCodeArr(j) = wsMaster.Cells(j, 3).Value
        masterCleanArr(j) = CleanText(masterRawArr(j))
        masterWordArr(j) = Split(masterCleanArr(j), " ")
    Next j

    Dim i As Long, outputRow As Long: outputRow = 2

    For i = 2 To lastMovieRow
        Dim rawMovie As String: rawMovie = wsMovie.Cells(i, 1).Value
        Dim movieTitle As String: movieTitle = ExtractQuotedTitle(rawMovie)
        Dim cleanedMovie As String: cleanedMovie = CleanText(movieTitle)
        Dim platformChannel As String: platformChannel = ExtractPlatformChannel(rawMovie)

        Dim movieWords() As String: movieWords = Split(cleanedMovie, " ")
        Dim movieCleanWords As Collection: Set movieCleanWords = RemoveStopWords(movieWords)

        Dim bestMatch As String, bestScore As Double: bestScore = 0
        Dim bestAssetType As String, bestAssetCode As String
        Dim tempMatches As Collection: Set tempMatches = New Collection

        For j = 2 To lastMasterRow
            Dim matchCount As Long: matchCount = 0
            Dim mw As Variant, k As Long
            For Each mw In movieCleanWords
                For k = LBound(masterWordArr(j)) To UBound(masterWordArr(j))
                    If StemWord(mw) = StemWord(masterWordArr(j)(k)) Then
                        matchCount = matchCount + 1
                        Exit For
                    End If
                Next k
            Next mw

            Dim currentScore As Double
            currentScore = matchCount / Application.WorksheetFunction.Max(1, movieCleanWords.count)

            If currentScore >= 0.5 And InStr(masterCleanArr(j), cleanedMovie) > 0 Then
                tempMatches.Add masterRawArr(j)
            End If

            If currentScore > bestScore Then
                bestScore = currentScore
                bestMatch = masterRawArr(j)
                bestAssetType = assetTypeArr(j)
                bestAssetCode = assetCodeArr(j)
            End If
        Next j

        Dim matchList As String: matchList = ""
        If tempMatches.count > 1 Then
            Dim item As Variant
            For Each item In tempMatches
                If matchList <> "" Then matchList = matchList & " , "
                matchList = matchList & item
            Next item
        End If

        Dim status As String, displayMatch As String
        If bestScore = 0 Then
            status = "No Match": displayMatch = ""
        ElseIf tempMatches.count > 1 Then
            status = "Needs Review": displayMatch = "Multiple matches found"
        ElseIf bestScore >= 0.75 Then
            status = "Matched": displayMatch = bestMatch
        Else
            status = "Low Match": displayMatch = bestMatch
        End If

        If status = "Matched" And matchList = "" Then matchList = "-"

        With wsOutput
            .Cells(outputRow, 1).Value = movieTitle
            .Cells(outputRow, 2).Value = status
            .Cells(outputRow, 3).Value = displayMatch
            .Cells(outputRow, 4).Value = Format(bestScore * 100, "0") & "%"
            .Cells(outputRow, 5).Value = IIf(status = "Matched", bestAssetType, "")
            .Cells(outputRow, 6).Value = IIf(status = "Matched", bestAssetCode, "")
            .Cells(outputRow, 7).Value = platformChannel
            .Cells(outputRow, 8).Value = matchList
        End With

        outputRow = outputRow + 1
    Next i

    ' ===================== CONDITIONAL FORMATTING =====================
    With wsOutput.Range("B2:B" & outputRow - 1)
        .FormatConditions.Delete
        .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Matched""").Interior.Color = RGB(144, 238, 144)
        .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Needs Review""").Interior.Color = RGB(255, 255, 102)
        .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Low Match""").Interior.Color = RGB(255, 204, 153)
        .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""No Match""").Interior.Color = RGB(255, 99, 71)
    End With

    ' Format enhancements
    wsOutput.Columns("A:H").AutoFit
    wsOutput.Range("A:H").WrapText = True
    wsOutput.Range("D2:D" & outputRow - 1).HorizontalAlignment = xlCenter

    ' Alternating row color
    Dim rowIndex As Long
    For rowIndex = 2 To outputRow - 1
        If rowIndex Mod 2 = 0 Then
            wsOutput.Range("A" & rowIndex & ":H" & rowIndex).Interior.Color = RGB(240, 248, 255)
        End If
    Next rowIndex

    ' Add borders
    With wsOutput.Range("A1:H" & outputRow - 1).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    ' Highlight placeholders
    Dim lastOutputRow As Long: lastOutputRow = outputRow - 1
    For i = 2 To lastOutputRow
        With wsOutput.Cells(i, 3)
            If .Value = "Multiple matches found" Then
                .Font.Bold = True
                .Font.Color = RGB(105, 105, 105)
                .HorizontalAlignment = xlCenter
            End If
        End With
        With wsOutput.Cells(i, 8)
            If .Value = "Already matched" Then
                .Font.Bold = True
                .Font.Color = RGB(105, 105, 105)
                .HorizontalAlignment = xlCenter
            End If
        End With
    Next i

    ' Freeze top row
    wsOutput.Activate
    wsOutput.Range("A2").Select
    ActiveWindow.FreezePanes = True

    ' Final touches
    wsOutput.Columns("A:H").AutoFit
    MsgBox "Matching complete!", vbInformation

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub
