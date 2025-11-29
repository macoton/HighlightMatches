Sub CreateSearchWordPuzzle()
Restart:
    Dim ws As Worksheet
    Set ws = Sheet2
    Dim gridSize As Integer: gridSize = 12
    Dim SearchWordCount As Integer: SearchWordCount = 15
    Dim maxWordLength As Integer: maxWordLength = 7
    Dim directions(1 To 8, 1 To 2) As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim startRow As Integer, startCol As Integer
    Dim wordLen As Integer
    Dim word As String
    Dim placed As Boolean
    Dim r As Integer, c As Integer
    Dim dirIndex As Integer
    Dim canPlace As Boolean
    Dim searchWordLengthCount As Integer
    Dim searchWordLength As Integer
    Dim searchWordLengths() As Integer
    Dim searchWordLengthIndex As Integer
    Dim searchWords() As String
    Dim wordIndex As Integer
    Dim occupied() As Boolean
    Dim attempt As Integer

    Randomize

    ' 8 directions: right, left, down, up, diag down-right, diag down-left, diag up-right, diag up-left
    directions(1, 1) = 0: directions(1, 2) = 1
    directions(2, 1) = 0: directions(2, 2) = -1
    directions(3, 1) = 1: directions(3, 2) = 0
    directions(4, 1) = -1: directions(4, 2) = 0
    directions(5, 1) = 1: directions(5, 2) = 1
    directions(6, 1) = 1: directions(6, 2) = -1
    directions(7, 1) = -1: directions(7, 2) = 1
    directions(8, 1) = -1: directions(8, 2) = -1

    ' Clear previous content
    ws.Range("A1:L12,N1:N15").ClearContents

    ' Initialize occupied array
    ReDim occupied(1 To gridSize, 1 To gridSize)
    For r = 1 To gridSize
        For c = 1 To gridSize
            occupied(r, c) = False
        Next c
    Next r

    ' Collect desired lengths from Sheet4!サーチワードの長さ
    ReDim searchWordLengths(1 To SearchWordCount)
    searchWordLengthCount = 0
    For i = 1 To SearchWordCount
        searchWordLength = Sheet4.Range("サーチワードの長さ").Cells(i, 1).Value
        If searchWordLength <> 0 Then
            searchWordLengthCount = searchWordLengthCount + 1
            searchWordLengths(searchWordLengthCount) = searchWordLength
        End If
    Next i

    ' Generate search words (random 2 to 7 digit numbers)
    ReDim searchWords(1 To SearchWordCount)
    For i = 1 To SearchWordCount
        searchWordLengthIndex = Int(searchWordLengthCount * Rnd) + 1 ' 1 to searchWordLengthCount
        wordLen = searchWordLengths(searchWordLengthIndex) ' 2 to 7
        ' consume the chosen length to avoid reuse
        For j = searchWordLengthIndex + 1 To searchWordLengthCount
            searchWordLengths(j - 1) = searchWordLengths(j)
        Next j
        searchWordLengths(searchWordLengthCount) = 0
        searchWordLengthCount = searchWordLengthCount - 1
        word = ""
        For j = 1 To wordLen
            word = word & CStr(Int(9 * Rnd) + 1) ' digits 1-9
        Next j
        searchWords(i) = word
    Next i

    ' Check for duplicates / substrings / reverse-substrings among searchWords
    Dim duplicateFound As Boolean
    Dim warnings As String
    Dim w1 As String, w2 As String, rw1 As String, rw2 As String

    duplicateFound = False
    warnings = ""

    For i = 1 To SearchWordCount - 1
        w1 = searchWords(i)
        rw1 = ReverseString(w1)
        For j = i + 1 To SearchWordCount
            w2 = searchWords(j)
            rw2 = ReverseString(w2)

            ' Exact duplicates
            If w1 = w2 Then
                duplicateFound = True
                warnings = warnings & "・同一: " & w1 & " vs " & w2 & " (行 " & i & " と " & j & ")" & vbCrLf
            End If

            ' Substring (either direction)
            If InStr(w2, w1) > 0 Or InStr(w1, w2) > 0 Then
                duplicateFound = True
                warnings = warnings & "・包含: " & w1 & " ? " & w2 & " (部分一致)" & vbCrLf
            End If

            ' Reverse-substring (either direction)
            If InStr(rw2, w1) > 0 Or InStr(rw1, w2) > 0 Then
                duplicateFound = True
                warnings = warnings & "・逆順包含: " & w1 & " ? rev(" & w2 & ")" & vbCrLf
            End If

            ' Exact reverse (full reversal equality)
            If w1 = rw2 Or w2 = rw1 Then
                duplicateFound = True
                warnings = warnings & "・逆順同一: " & w1 & " = rev(" & w2 & ")" & vbCrLf
            End If
        Next j
    Next i

    If duplicateFound Then
'        MsgBox "難易度低下の恐れあり：サーチワード間で重複・包含・逆順包含が見つかりました。" & vbCrLf & vbCrLf & warnings, vbExclamation
        GoTo Restart
    End If

    ' Place search words in the grid
    For wordIndex = 1 To SearchWordCount
        word = searchWords(wordIndex)
        placed = False
        For attempt = 1 To 100
            startRow = Int(gridSize * Rnd) + 1
            startCol = Int(gridSize * Rnd) + 1
            dirIndex = Int(8 * Rnd) + 1
            canPlace = True
            ' Check if word fits
            For k = 0 To Len(word) - 1
                r = startRow + directions(dirIndex, 1) * k
                c = startCol + directions(dirIndex, 2) * k
                If r < 1 Or r > gridSize Or c < 1 Or c > gridSize Then
                    canPlace = False
                    Exit For
                End If
                If occupied(r, c) Then
                    ' Allow overlap only if same digit
                    If ws.Cells(r, c).Value <> Mid(word, k + 1, 1) Then
                        canPlace = False
                        Exit For
                    End If
                End If
            Next k
            If canPlace Then
                ' Place the word
                For k = 0 To Len(word) - 1
                    r = startRow + directions(dirIndex, 1) * k
                    c = startCol + directions(dirIndex, 2) * k
                    ws.Cells(r, c).Value = Mid(word, k + 1, 1)
                    occupied(r, c) = True
                Next k
                placed = True
                Exit For
            End If
        Next attempt
        If Not placed Then
'            MsgBox "Failed to place word: " & word
            GoTo Restart
        End If
    Next wordIndex

    ' Fill remaining empty cells with random digits 1-9
    For r = 1 To gridSize
        For c = 1 To gridSize
            If ws.Cells(r, c).Value = "" Then
                ws.Cells(r, c).Value = Int(9 * Rnd) + 1
            End If
        Next c
    Next r

    ' Output search words to N1:N15
    For i = 1 To SearchWordCount
        ws.Cells(i, 14).Value = searchWords(i)
    Next i

    ws.Range("N1:N15").Font.Bold = False
    ws.Range("N1:N15").Font.Color = vbBlack
    ws.Range("N1:N15").Font.Size = 11

    MsgBox "Puzzle created successfully!"
End Sub

Private Function ReverseString(ByVal s As String) As String
    Dim i As Long, t As String
    t = ""
    For i = Len(s) To 1 Step -1
        t = t & Mid$(s, i, 1)
    Next i
    ReverseString = t
End Function
