Sub BingoCardDistribution()
'カード配布
Dim i As Long, j As Long, k As Long
Dim x()
Dim y()
Dim d, d2(1 To 5, 1 To 5)
Dim ld, ud, pd
Dim cd(1 To 24) As String

    '---カードの位置指定
    cd(1) = "C3": cd(2) = "J3": cd(3) = "C10": cd(4) = "J10": cd(5) = "C17": cd(6) = "J17": cd(7) = "C24": cd(8) = "J24"
    cd(9) = "Q3": cd(10) = "X3": cd(11) = "Q10": cd(12) = "X10": cd(13) = "Q17": cd(14) = "X17": cd(15) = "Q24": cd(16) = "X24"
    cd(17) = "AE3": cd(18) = "AL3": cd(19) = "AE10": cd(20) = "AL10": cd(21) = "AE17": cd(22) = "AL17": cd(23) = "AE24": cd(24) = "AL24"
    
    '---シートのクリア
    Worksheets("Sheet2").Range("B:E").ClearContents
    Worksheets("Sheet1").Range("C:AP").Interior.ColorIndex = xlNone
    Worksheets("Sheet1").Range("C:AP").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Worksheets("Sheet1").Range("A1").Value = 0
    Worksheets("Sheet1").Range("A2").Value = ""
    'Worksheets("Sheet1").Range("C1,E1,J1,L1,C8,E8,J8,L8").Value = ""
    Worksheets("Sheet1").Range("C:AP").Value = ""
    With Worksheets("Sheet1")
        For k = 1 To 24    'カードの枚数
            For j = 1 To 5    '列ごとに乱数を取り出す
                '---乱数の開始値Ldと終了値Udを指定
                ld = 1 + (j - 1) * 15     '----開始値
                ud = 15 * j     '----終了値
                pd = 5     '----取り出す個数
                '----使用する配列を準備する(1列分)
                ReDim x(1 To ud - ld + 1)
                ReDim y(1 To ud - ld + 1)
                ReDim d(1 To ud - ld + 1, 1 To 1)
                Randomize
                '----乱数と値を配列にセットする
                For i = 1 To ud - ld + 1
                    x(i) = Rnd()
                    y(i) = i + ld - 1
                Next i
                '----値を取り出す(1列分の5個の数値)
                For i = 1 To ud - ld + 1
                    d(i, 1) = y(Application.Match(Application.Small(x, i), x, 0))
                Next i
                '----カード1枚分の数値を配列に収める
                For i = 1 To 5
                    d2(i, j) = d(i, 1)
                Next i
            Next j
            '----シートにカードを書き出す
            .Range(cd(k)).Resize(5, 5).Value = d2
            .Range(cd(k)).Offset(2, 2).Value = "F"  '中央を「Free」とする
            .Range(cd(k)).Offset(2, 2).Interior.ColorIndex = 6    '中央を黄色で塗りつぶす

        Next k
    End With
End Sub
Sub NumberCreate()
    '抽選番号を作成する
    Dim i As Long, j As Long
    Dim x()
    Dim y()
    Dim d
    '---開始値Ldと終了値Udを定数で指定
    Const ld = 1
    Const ud = 75
    '----使用する配列を準備する
    ReDim x(1 To ud - ld + 1)
    ReDim y(1 To ud - ld + 1)
    ReDim d(1 To ud - ld + 1, 1 To 1)
    Randomize

    With Worksheets("Sheet2")
        .Range("G:G").ClearContents
        '----乱数と値を配列にセットする
        For i = 1 To ud - ld + 1
            x(i) = Rnd()
            y(i) = i + ld - 1
        Next i
        '----値を取り出す
        For i = 1 To 75
            d(i, 1) = y(Application.Match(Application.Small(x, i), x, 0))
        Next i
        .Cells(1, 7).Resize(UBound(d), 1) = d
    End With
End Sub
Sub NumberCollation()
    '番号を照会する
    Dim i As Long, j As Long, k As Long
    Dim cd(1 To 24) As String
    '----ビンゴカードの左上の位置を指定しています
    cd(1) = "C3": cd(2) = "J3": cd(3) = "C10": cd(4) = "J10": cd(5) = "C17": cd(6) = "J17": cd(7) = "C24": cd(8) = "J24"
    cd(9) = "Q3": cd(10) = "X3": cd(11) = "Q10": cd(12) = "X10": cd(13) = "Q17": cd(14) = "X17": cd(15) = "Q24": cd(16) = "X24"
    cd(17) = "AE3": cd(18) = "AL3": cd(19) = "AE10": cd(20) = "AL10": cd(21) = "AE17": cd(22) = "AL17": cd(23) = "AE24": cd(24) = "AL24"

    With Worksheets("Sheet1")
        .Range("A1").Value = .Range("A1").Value + 1
        .Range("A2").Value = Worksheets("Sheet2").Range("G" & .Range("A1").Value).Value

        For k = 1 To 24
            For i = 1 To 5
                For j = 1 To 5
                    If .Range(cd(k)).Offset(i - 1, j - 1).Value = .Range("A2").Value Then
                        .Range(cd(k)).Offset(i - 1, j - 1).Interior.ColorIndex = 6
                    End If
                Next j
            Next i
        Next k
    End With
    NumberCheck
End Sub
Sub NumberCheck()
    'カードCheck
    Dim i As Long, j As Long, k As Long
    Dim cn1 As Integer, cn2 As Integer, cn3 As Integer, cn4 As Integer
    Dim cd(1 To 24) As String, ce(1 To 24) As String, cf(1 To 24) As String
    Dim y()
    Dim d
        '----カードの位置とリーチ、ビンゴの表示位置
        cd(1) = "C3": cd(2) = "J3": cd(3) = "Q3": cd(4) = "X3": cd(5) = "AE3": cd(6) = "AL3"
        cd(7) = "C10": cd(8) = "J10": cd(9) = "Q10": cd(10) = "X10": cd(11) = "AE10": cd(12) = "AL10"
        cd(13) = "C17": cd(14) = "J17": cd(15) = "Q17": cd(16) = "X17": cd(17) = "AE17": cd(18) = "AL17"
        cd(19) = "C24": cd(20) = "J24": cd(21) = "Q24": cd(22) = "X24": cd(23) = "AE24": cd(24) = "AL24"
        
        ce(1) = "D2": ce(2) = "K2": ce(3) = "R2": ce(4) = "Y2": ce(5) = "AF2": ce(6) = "AM2"
        ce(7) = "D9": ce(8) = "K9": ce(9) = "R9": ce(10) = "Y9": ce(11) = "AF9": ce(12) = "AM9"
        ce(13) = "D16": ce(14) = "K16": ce(15) = "R16": ce(16) = "Y16": ce(17) = "AF16": ce(18) = "AM16"
        ce(19) = "D23": ce(20) = "K23": ce(21) = "R23": ce(22) = "Y23": ce(23) = "AF23": ce(24) = "AM23"
        
        cf(1) = "F2": cf(2) = "M2": cf(3) = "T2": cf(4) = "AA2": cf(5) = "AH2": cf(6) = "AO2"
        cf(7) = "F9": cf(8) = "M9": cf(9) = "T9": cf(10) = "AA9": cf(11) = "AH9": cf(12) = "AO9"
        cf(13) = "F16": cf(14) = "M16": cf(15) = "T16": cf(16) = "AA16": cf(17) = "AH16": cf(18) = "AO16"
        cf(19) = "F23": cf(20) = "M23": cf(21) = "T23": cf(22) = "AA23": cf(23) = "AH23": cf(24) = "AO23"
        '---縦と横方向のチェック
      With Worksheets("Sheet1")
          For k = 1 To 24
              cn1 = 0: cn2 = 0: cn3 = 0: cn4 = 0
              For i = 1 To 5
                  For j = 1 To 5
                      If .Range(cd(k)).Offset(i - 1, j - 1).Interior.ColorIndex = 6 Then
                          cn1 = cn1 + 1
                      End If

                      If .Range(cd(k)).Offset(j - 1, i - 1).Interior.ColorIndex = 6 Then
                          cn2 = cn2 + 1
                      End If
                  Next j
                  '----リーチとビンゴの判定
                  If cn1 = 4 Or cn2 = 4 Then
                      .Range(ce(k)).Value = "リーチ"
                  End If
                  If cn1 = 5 Or cn2 = 5 Then
                      .Range(cf(k)).Value = "BINGO!!"
                  End If
                  cn1 = 0: cn2 = 0
              Next i
        '----斜め方向をチェック
        For i = 1 To 5
            If .Range(cd(k)).Offset(i - 1, i - 1).Interior.ColorIndex = 6 Then
                cn3 = cn3 + 1
            End If
            If .Range(cd(k)).Offset(i - 1, 5 - i).Interior.ColorIndex = 6 Then
                cn4 = cn4 + 1
            End If
        Next i
        '----リーチとビンゴの判定
        If cn3 = 4 Or cn4 = 4 Then
            .Range(ce(k)).Value = "リーチ"
        End If
        If cn3 = 5 Or cn4 = 5 Then
            .Range(cf(k)).Value = "BINGO!!"
        End If
        cn3 = 0: cn4 = 0

      Next k

  End With
End Sub
