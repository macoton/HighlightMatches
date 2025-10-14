Option Explicit

Sub HighlightMatches()
    Dim grid(1 To 12, 1 To 12) As String
    Dim targets(1 To 15) As String
    Dim i As Integer, j As Integer, k As Integer
    Dim dirX As Variant, dirY As Variant
    Dim t As Integer, r As Integer
    Dim match As Boolean
    
    ' 方向ベクトル（8方向）
    dirX = Array(0, 1, 1, 1, 0, -1, -1, -1)
    dirY = Array(1, 1, 0, -1, -1, -1, 0, 1)
    
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
    For t = 1 To 15
        Dim target As String
        target = targets(t)
        
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
                        ' 色付け（例：黄色）
                        For r = 1 To Len(target)
                            posX = i + dirX(k) * (r - 1)
                            posY = j + dirY(k) * (r - 1)
                            Cells(posX, posY).Interior.Color = RGB(255, 255, 0)
                        Next r
                    End If
                Next k
            Next j
        Next i
    Next t
End Sub

Sub ResetColors()
    Dim i As Integer, j As Integer
    ' 12×12の範囲の色をクリア
    For i = 1 To 12
        For j = 1 To 12
            Cells(i, j).Interior.ColorIndex = xlColorIndexNone
        Next j
    Next i
End Sub
