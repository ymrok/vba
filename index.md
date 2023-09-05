# 3 日目の復習 ＋ α

## セル内の文字（数字）に色を付ける

### 色コードを使用する

セル内の文字を色コードで着色します。

| 色コード | 文字の色 |
| :---: | :---: |
| vbBlack | 黒 |
| vbRed | 赤 |
| vbGreen | 緑 |
| vbYellow | 黄 |
| vbBlue | 青 |
| vbMagenta | マゼンタ |
| vbCyan | シアン |
| vbWhite | 白 |

セルは範囲指定も可能です。

```vb
Range(セル位置).Font.Color = 色コード
```

シート「データ」のセル A1 の文字を赤色にするコードです。

```vb
Private Sub CommandButton1_Click()

    ' セル内の文字に色を付ける
    Worksheets("データ").Range("A1").Font.Color = vbRed     ' 赤色

End Sub
```

実行前の状態です。

![実行前](img/2023-09-05_18h07_47.png)

![実行前](img/2023-09-05_18h08_55.png)

実行後のシート「データ」の状態です。セル A1 の文字が赤色になりました。

![実行後](img/2023-09-05_18h12_34.png)

### RGB コードで指定した色を付ける

セル内の文字を RGB で表現した値で着色します。各色は 0 ～ 255 までのいずれかの値を取ります。色のサンプルは下記のようなサイトでご確認ください。

現色大辞典  
<https://www.colordic.org/>

セル位置は範囲しても可能です。

```vb
Range(セル位置).Font.Color = RGB(赤, 緑, 青)
```

シート「データ」のセル A2 ～ B2 の文字をエンジ色にするコードです。

```vb
Private Sub CommandButton1_Click()

    ' セル内の文字に色を付ける
    Worksheets("データ").Range("A2", "B2").Font.Color = RGB(229, 23, 31)    ' エンジ色

End Sub
```

実行前の状態です。

![実行前](img/2023-09-05_18h07_47.png)

![実行前](img/2023-09-05_18h08_55.png)

実行後のシート「データ」の状態です。セル A2 ～ B2 の数字がエンジ色になりました。

![実行後](img/2023-09-05_18h20_42.png)

## 罫線を引く

### 引く罫線の種類を指定する

#### LineStyle で指定

LinsStyele で指定する場合、線の太さは指定できません。

| 線の種類 | 引かれる線 |
| :--- | :--- |
| xlLineStyleNon | 線なし |
| xlContinuous | 実線 |
| xlDot | 点線 |
| xlDashDotDot | 二点鎖線 |
| xlDashDot | 一点鎖線 |
| xlDash | 破線 |
| xlDouble | 二本線 |

#### Weight で指定

線の太さを指定する場合、引かれる線は実線だけです。

| 線の種類 | 引かれる線 |
| :--- | :--- |
| xlMedium | 中太線 |
| xlThick | 太線 |

### 指定したセルの範囲の外枠に罫線を引く

```vb
Range(セル位置).BorderAround LineStyle := 線の種類
```

```vb
Range(セル位置).BorderAround Weight := 線の種類
```

### 指定したセルの範囲内に格子の罫線を引く

```vb
Range(セル位置).Borders.LineStyle = 線の種類
```

```vb
Range(セル位置).Borders.Weight = 線の種類
```

### 罫線を引くときに注意すること

罫線は後から引いたほうが優先されます。

1. 破線の罫線を引く
2. 実線の罫線を引く

この場合、後から引いた実線だけが残ります。罫線が重なり合う部分がある場合、どの順番で引くのか考慮しなければなりません。

シート「データ」の値に罫線を追加するコードです。

```vb
Private Sub CommandButton1_Click()

    ' 罫線を引く
    Worksheets("データ").Range("A1", "B7").Borders.LineStyle = xlContinuous     ' 実線で格子
    Worksheets("データ").Range("A1", "B7").BorderAround Weight:=xlThick         ' 外枠は太枠

End Sub
```

実行前の状態です。

![実行前](img/2023-09-05_18h07_47.png)

![実行前](img/2023-09-05_20h33_59.png)

実行後のシート「データ」の状態です。格子状の罫線と外側は太線で囲まれています。

![実行後](img/2023-09-05_20h37_45.png)

## With

次のコードは `Worksheets("データ").Range("A1", "B7")` が 2 回使用されています。このように同一のオブジェクトが複数回使用される場合、`With` でまとめて記述できます。

```vb
Private Sub CommandButton1_Click()

    ' 罫線を引く
    .Borders.LineStyle = xlContinuous     ' 実線で格子
    Worksheets("データ").Range("A1", "B7").BorderAround Weight:=xlThick         ' 外枠は太枠

End Sub
```

`With` の使い方です。

```vb
With オブジェクト

    .VBA のコード

End With
```

`With` ･･･ `End With` で囲まれた VBA のコードは `With` の適用範囲がわかるように（明示）するためインデント（字下げ）して記述します。

上述のコードを `With` で書き直したコードです。

```vb
Private Sub CommandButton1_Click()

    ' 罫線を引く
    With Worksheets("データ").Range("A1", "B7")
        .Borders.LineStyle = xlContinuous                   ' 実線で格子
        .BorderAround Weight:=xlThick                       ' 外枠は太枠
    End With
    
End Sub
```

`With` ･･･ `End With` でくくられた VBA コードで、先頭が `.` で始まるコードだけに `With` の右横に記述したオブジェクト（上記の例では `Worksheets("データ").Range("A1", "B7")` ）が自動的に補われて実行します。

`With` の右横にオブジェクト型の値を定義するため、次のように記述できます。

```vb
Private Sub CommandButton1_Click()

    Dim Range_Table     As Range                            ' 表の範囲
   
    ' テーブルの範囲を保存
    Set Range_Table = Worksheets("データ").Range("A1", "B7")
    
    ' 罫線を引く
    With Range_Table
        .Borders.LineStyle = xlContinuous                   ' 実線で格子
        .BorderAround Weight:=xlThick                       ' 外枠は太枠
    End With
    
End Sub
```

## Range.CurrentRegion

```vb
Range(基準になるセル位置).CurrentRegion
```

`Range` でセルの範囲指定を行うときに `CurrentRegion` を使用すると、範囲指定が柔軟に行えることがあります。 `CurrentRegion` は基準となるセルを指定し、そのセルに隣接していて　かつ　値が入っているセルをひとまとめで指定できます。

![実行前](img/2023-09-05_20h33_59.png)

基準セルを A1 にして `CurrentRegion` を使用すると、黄色で塗りつぶした範囲が対象になります。

![黄色](img/2023-09-05_21h25_33.png)

先程の罫線を引くコードを `CurrentRegion` を使用して書き直しました。

```vb
Private Sub CommandButton1_Click()

    Dim Range_Table     As Range                            ' 表の範囲
   
    ' テーブルの範囲を保存
    Set Range_Table = Worksheets("データ").Range("A1").CurrentRegion
    
    ' 罫線を引く
    With Range_Table
        .Borders.LineStyle = xlContinuous                   ' 実線で格子
        .BorderAround Weight:=xlThick                       ' 外枠は太枠
    End With
    
End Sub
```

実行前のシート「データ」の状態です。

![実行前](img/2023-09-05_20h33_59.png)

実行後のシート「データ」の状態です。正しく罫線を引きました。

![実行後](img/2023-09-05_20h37_45.png)

`CurrentRegion` が便利なのは、データが増えた場合でもコードの変更が不要なことです。

シート「データ」が次のように書き足されました。

![実行前](img/2023-09-05_21h35_25.png)

再度の実行後のシート「データ」の状態です。正しく罫線を引き直しました。

![実行後](img/2023-09-05_21h36_52.png)
