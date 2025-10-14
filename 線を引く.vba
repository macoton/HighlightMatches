Option Explicit

Sub HighlightAllMatches()
    HighlightMatches True
End Sub

Sub HighlightSelectMatches()
    HighlightMatches False
End Sub

Sub ResetColors()
    Dim i As Integer, j As Integer
    ' 12×12の範囲の色をクリア
    For i = 1 To 12
        For j = 1 To 12
            Cells(i, j).Interior.ColorIndex = xlColorIndexNone
        Next j
    Next i

    Dim sh As Shape
    For Each sh In ActiveSheet.Shapes
        If Left(sh.Name, 5) = "Match" Then sh.Delete
    Next sh

End Sub

Sub HighlightMatches(All As Boolean)
    Dim grid(1 To 12, 1 To 12) As String
    Dim targets(1 To 15) As String
    Dim i As Integer, j As Integer, k As Integer
    Dim dirX As Variant, dirY As Variant
    Dim t As Integer, r As Integer
    Dim match As Boolean
    Dim offset As Integer
    Dim matchCount As Integer
    Dim matchCount2(1 To 12, 1 To 12) As Integer

    ' 方向ベクトル（8方向）
    dirX = Array(0, 1, 1, 1, 0, -1, -1, -1)
    dirY = Array(1, 1, 0, -1, -1, -1, 0, 1)

    ResetColors

    ' マス目読み込み
    For i = 1 To 12
        For j = 1 To 12
            grid(i, j) = Cells(i, j).Value
        Next j
    Next i

    ' 探す数値読み込み
    For i = 1 To 15
        targets(i) = Cells(i, 14).Value ' N列
    Next i

    ' 探索処理
    matchCount = 0
    For i = 1 To 12
        For j = 1 To 12
            matchCount2(i, j) = 0
        Next j
    Next i
    For t = 1 To 15
        Dim target As String
        target = targets(t)

        If All Or _
            Not Intersect(Selection, Cells(t, 14)) Is Nothing Then
            For i = 1 To 12
                For j = 1 To 12
                    For k = 0 To 7 ' 8方向
                        match = True
                        Dim posX As Integer, posY As Integer
    
                        For r = 1 To Len(target)
                            posX = i + dirX(k) * (r - 1)
                            posY = j + dirY(k) * (r - 1)
    
                            If posX < 1 Or posX > 12 Or posY < 1 Or posY > 12 Then
                                match = False
                                Exit For
                            End If
    
                            If Mid(target, r, 1) <> grid(posX, posY) Then
                                match = False
                                Exit For
                            End If
                        Next r
    
                        If match Then
                            If Range("セルを黄色").Value Then
                                ' 色付け（例：黄色）
                                For r = 1 To Len(target)
                                    posX = i + dirX(k) * (r - 1)
                                    posY = j + dirY(k) * (r - 1)
                                    Cells(posX, posY).Interior.Color = RGB(255, 255, 0)
                                Next r
                            End If
                            If Range("線を引く").Value Then
                                offset = matchCount2(i, j) * 2 ' Shift each line slightly
                                Call DrawMatchLine(i, j, CInt(dirX(k)), CInt(dirY(k)), Len(target), offset, "Match" & matchCount)
                            End If
                            matchCount = matchCount + 1
                            matchCount2(i, j) = matchCount2(i, j) + 1
                        End If
                    Next k
                Next j
            Next i
        End If
    Next t

End Sub

Sub DrawMatchLine(startRow As Integer, startCol As Integer, dirX As Integer, dirY As Integer, length As Integer, offset As Integer, shapeName As String)
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim cell1 As Range, cell2 As Range

    Set cell1 = ws.Cells(startRow, startCol)
    Set cell2 = ws.Cells(startRow + dirX * (length - 1), startCol + dirY * (length - 1))

    x1 = cell1.Left + cell1.Width / 2 + offset
    y1 = cell1.Top + cell1.Height / 2 + offset
    x2 = cell2.Left + cell2.Width / 2 + offset
    y2 = cell2.Top + cell2.Height / 2 + offset

    With ws.Shapes.AddLine(x1, y1, x2, y2)
        .Name = shapeName
        .Line.ForeColor.RGB = RGB(255 - offset * 10, 100 + offset * 5, 150 + offset * 5)
        .Line.Weight = 2
    End With

End Sub
