Attribute VB_Name = "M_Macros"
'******************************
'Excel Utility Macros
'******************************

Option Explicit

#If VBA7 And Win64 Then
Private Declare PtrSafe Function OpenClipboard Lib "user32.dll" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32.dll" () As Long
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As LongLong) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32.dll" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalSize Lib "kernel32.dll" (ByVal hMem As LongPtr) As LongLong
Private Declare PtrSafe Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal length As LongLong)
#Else
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
#End If

Public Sub SetClipboard(sUniText As String)
#If VBA7 And Win64 Then
    Dim iStrPtr As LongPtr
    Dim iLen As LongLong
    Dim iLock As LongPtr
#Else
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
#End If
    Const GMEM_MOVEABLE As Long = &H2
    Const GMEM_ZEROINIT As Long = &H40
    Const CF_UNICODETEXT As Long = &HD
    OpenClipboard 0&
    EmptyClipboard
    iLen = LenB(sUniText)
    iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen + 2&)
    iLock = GlobalLock(iStrPtr)
    MoveMemory iLock, StrPtr(sUniText), iLen
    GlobalUnlock iStrPtr
    SetClipboardData CF_UNICODETEXT, iStrPtr
    CloseClipboard
End Sub

Public Function GetClipboard() As String
#If VBA7 And Win64 Then
    Dim iStrPtr As LongPtr
    Dim iLen As LongLong
    Dim iLock As LongPtr
#Else
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
#End If
    Dim sUniText As String
    Const CF_UNICODETEXT As Long = 13&
    OpenClipboard 0&
    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        iStrPtr = GetClipboardData(CF_UNICODETEXT)
        If iStrPtr Then
            iLock = GlobalLock(iStrPtr)
            iLen = GlobalSize(iStrPtr)
            sUniText = String$(CLng(iLen) \ 2& - 1&, vbNullChar)
            MoveMemory StrPtr(sUniText), iLock, LenB(sUniText)
            GlobalUnlock iStrPtr
        End If
        GetClipboard = sUniText
    End If
    CloseClipboard
End Function


Sub CopyStringToClipboard()

    SetClipboard "Hello, 私は人間です"
    
End Sub


'******************************
'シートの選択ダイアログを表示する
'引数  ：なし
'戻り値：なし
'******************************
Sub シート選択Dialog_EUM()
   
   With CommandBars.Add(Temporary:=True)
       .Controls.Add(ID:=957).Execute
       .Delete
   End With
   
End Sub

'******************************
'数式バー、ヘッダー、スクロールバー表示切り替え
'引数  ：なし
'戻り値：なし
'******************************
Sub 数式ヘッダスクロール表示切替_EUM()

    ' 数式バーの表示状態を切り替える
    Application.DisplayFormulaBar = Not Application.DisplayFormulaBar
    With ActiveWindow
        If Application.DisplayFormulaBar Then
            ' スクロールバーを表示
            .DisplayHorizontalScrollBar = True
            .DisplayVerticalScrollBar = True
            ' ヘッダーを表示
            .DisplayHeadings = True
        Else
            ' スクロールバーを非表示
            .DisplayHorizontalScrollBar = False
            .DisplayVerticalScrollBar = False
            ' ヘッダーを非表示
            .DisplayHeadings = False
        End If
        ShowToast "表示／非表示を切り替えました" & vbCrLf & "[scroll][formula][col&row head]"
    End With
End Sub

'******************************
'全画面表示
'引数  ：なし
'戻り値：なし
'******************************
Sub 全画面表示_EUM()
    With ActiveWindow
        If Application.DisplayFullScreen Then
        Else
            If Application.DisplayFormulaBar Then
                ' スクロールバーを表示
                .DisplayHorizontalScrollBar = True
                .DisplayVerticalScrollBar = True
                ' ヘッダーを表示
                .DisplayHeadings = True
                ' 全画面表示モードに設定
                Application.DisplayFullScreen = True
                ' 数式バーの表示状態を表示
                Application.DisplayFormulaBar = True
            Else
                ' スクロールバーを非表示
                .DisplayHorizontalScrollBar = False
                .DisplayVerticalScrollBar = False
                ' ヘッダーを非表示
                .DisplayHeadings = False
                ' 全画面表示モードに設定
                Application.DisplayFullScreen = True
                ' 数式バーの表示状態を非表示
                Application.DisplayFormulaBar = False
            End If
        End If
    End With
    ShowToast "全画面終了するにはEcsキー"
End Sub

'******************************
'アクティブブックの名前をクリップボードにコピー
'引数  ：なし
'戻り値：なし
'******************************
Sub ブック名コピー_EUM()
    Dim bookName As String
    
    ' アクティブブックの名前を取得
    bookName = ActiveWorkbook.Name
    
    ' クリップボードにブック名をコピーする
    If Len(bookName) > 0 Then
        SetClipboard bookName
        ShowToast "ブック名がクリップボードにコピーされました。"
    Else
        MsgBox "ブック名が取得できませんでした。"
    End If
End Sub

'******************************
'アクティブブックのフルパスをクリップボードにコピー
'引数  ：なし
'戻り値：なし
'******************************
Sub ブックパスコピー_EUM()
    Dim bookPath As String
    
    ' アクティブブックのフルパスを取得
    bookPath = ActiveWorkbook.FullName
    
    ' クリップボードにブックのフルパスをコピーする
    If Len(bookPath) > 0 Then
        SetClipboard bookPath
        ShowToast "ブックのフルパスがクリップボードにコピーされました。"
    Else
        MsgBox "ブックのフルパスが取得できませんでした。"
    End If
End Sub

'******************************
'アクティブブックのパスをエクスプローラーで表示
'引数  ：なし
'戻り値：なし
'******************************
Sub ブックパスをエクスプローラーで表示_EUM()
    Dim bookPath As String
    
    ' アクティブブックのフルパスを取得
    bookPath = ActiveWorkbook.FullName
    
    ' ブックのパスが取得できた場合はエクスプローラーで表示する
    If Len(bookPath) > 0 Then
        Shell "explorer.exe /select," & bookPath, vbNormalFocus
    Else
        MsgBox "ブックのパスが取得できませんでした。"
    End If
End Sub


'******************************
'選択セルの内容の名前でシート作成
'引数  ：なし
'戻り値：なし
'******************************
Sub シート作成SelectCell2SheetName_EUM()
    Dim baseSheetName As String
    Dim sheetName As String
    Dim sheetNumber As Integer
    Dim newSheet As Worksheet
    Dim originalSheet As Worksheet
    Dim cell As Range

    ' 現在のアクティブシートを保存
    Set originalSheet = ActiveSheet

    ' 選択範囲内の各セルに対して処理
    For Each cell In Selection
        baseSheetName = ReplaceInvalidCharacters(cell.Text)

        ' 最長31文字に制限
        If Len(baseSheetName) > 31 Then
            baseSheetName = Left(baseSheetName, 31)
        End If

        sheetNumber = 1
        sheetName = baseSheetName

        ' シート名が既に存在する場合は番号を増やす
        While SheetExistsInActiveWorkbook(sheetName)
            sheetNumber = sheetNumber + 1
            sheetName = baseSheetName & "(" & sheetNumber & ")"
        Wend

        ' 新しいシートを作成
        Set newSheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        
        If sheetName <> "" Then newSheet.Name = sheetName
        
        ' A1にシート名
        newSheet.Range("A1").Formula = "=MID(@CELL(""filename"",A1),FIND(""]"",@CELL(""filename"",A1))+1,31)"

        ' A2に数式を設定
        newSheet.Range("A2").Formula = "=IFERROR(HYPERLINK(""#'" & """ & B2 & ""'!A2"", "" << ""), """")"
        With newSheet.Range("A2").Font
            .Name = "メイリオ"  ' フォントをメイリオに設定
            .Color = RGB(0, 112, 192)  ' 文字色をHEX #0070C0（RGBで0, 112, 192）に設定
            .Underline = xlUnderlineStyleSingle  ' 下線を設定
        End With

        ' B2に元のシート名を設定
        newSheet.Range("B2").value = originalSheet.Name
        
    Next cell

    ' 元のシートに戻る
    originalSheet.Activate
End Sub

' シート名に使用できない文字を除外する関数
Function ReplaceInvalidCharacters(inputString As String) As String
    Dim invalidChars As Variant
    invalidChars = Array("/", "\", ":", "*", "?", """", "<", ">", "|", "・", Chr(10), Chr(13))

    Dim i As Integer
    For i = LBound(invalidChars) To UBound(invalidChars)
        inputString = Replace(inputString, invalidChars(i), "")
    Next i

    ReplaceInvalidCharacters = inputString
End Function

' アクティブなワークブック内でシートが存在するかどうかを確認するヘルパー関数
Function SheetExistsInActiveWorkbook(sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ActiveWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExistsInActiveWorkbook = Not sheet Is Nothing
End Function

'******************************
'全てのシートでA1選択
'引数  ：なし
'戻り値：なし
'******************************
Sub 全てのシートでA1選択_EUM()
    Dim ws As Worksheet
    Dim firstSheet As Worksheet

    ' ブック内の先頭シートを取得
    Set firstSheet = ActiveWorkbook.Sheets(1)

    ' アプリケーションのスクリーンアップデートを停止する（処理を高速化）
    Application.ScreenUpdating = False

    ' すべてのシートをループで処理
    For Each ws In ActiveWorkbook.Sheets
        ' シートをアクティブにしてA1セルをアクティブにする
        ws.Activate
        ws.Range("A1").Activate
        
        ' 表示倍率を80%に設定
        ActiveWindow.Zoom = 80
        
        ' 表示域を左上端へスクロール
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    Next ws

    ' アプリケーションのスクリーンアップデートを再開する
    Application.ScreenUpdating = True

    ' ブック内の先頭シートをアクティブにして表示
    firstSheet.Activate

    ShowToast "全シートA1セル選択、表示倍率：80%"
End Sub


'******************************
'すべてのシート名をクリップボードにコピー
'引数  ：なし
'戻り値：なし
'******************************
Sub シート名コピーALL_EUM()
    Dim ws As Worksheet
    Dim strSheetNames As String
    
    'すべてのシート名を連結した文字列を作成する
    For Each ws In ActiveWorkbook.Worksheets
        strSheetNames = strSheetNames & ws.Name & vbCrLf
    Next
    
    ' クリップボードにシート名をコピーする
    If Len(strSheetNames) > 0 Then
        SetClipboard strSheetNames
        ShowToast "すべてのシート名がクリップボードにコピーされました。"
    Else
        MsgBox "このブックにはシートがありません。"
    End If
End Sub

'******************************
'選択中のシート名をクリップボードにコピー
'引数  ：なし
'戻り値：なし
'******************************
Sub シート名コピーSelected_EUM()
    Dim ws As Worksheet
    Dim strSheetNames As String

    ' 選択中のシート名を連結した文字列を作成する
    For Each ws In ActiveWindow.SelectedSheets
        strSheetNames = strSheetNames & ws.Name & vbCrLf
    Next

    ' クリップボードにシート名をコピーする
    If Len(strSheetNames) > 0 Then
        SetClipboard strSheetNames
        ShowToast "選択されたシート名がクリップボードにコピーされました。"
    Else
        MsgBox "選択されたシートがありません。"
    End If
End Sub

'******************************
'選択中オブジェクトを最背面へ
'引数  ：なし
'戻り値：なし
'******************************
Sub オブジェクトを最背面へ_EUM()
    On Error GoTo ErrorHandler
    Selection.ShapeRange.ZOrder msoSendToBack
    Exit Sub
ErrorHandler:
    Application.Wait Now + TimeValue("0:00:01") ' 1秒待機
End Sub

'******************************
'選択中オブジェクトを最前面へ
'引数  ：なし
'戻り値：なし
'******************************
Sub オブジェクトを最前面へ_EUM()
    On Error GoTo ErrorHandler
    Selection.ShapeRange.ZOrder msoBringToFront
    Exit Sub
ErrorHandler:
    Application.Wait Now + TimeValue("0:00:01") ' 1秒待機
End Sub

'******************************
'吹き出し図形を挿入する
'引数  ：なし
'戻り値：なし
'******************************
Sub 吹き出し挿入_EUM()
    Dim ws As Worksheet
    Dim shp As Shape
    
    ' アクティブシートを設定
    Set ws = ActiveSheet
    
    ' 吹き出し図形を挿入
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangularCallout, 100, 100, 150, 100)
    
    With shp
        ' テキストを設定
        .TextFrame.characters.Text = "*"
        
        ' フォント設定
        With .TextFrame2.TextRange.Font
            .Name = "Arial"
            .Size = 24
            .Fill.ForeColor.RGB = RGB(0, 0, 0)  ' 黒色
        End With
        
        ' 図形の塗りつぶし
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 200)  ' 淡い黄色
            .Transparency = 0.2  ' 20%の透明度
        End With
        
        ' 枠線の設定
        With .Line
            .Visible = msoTrue
            .Weight = 2.25  ' 太さ2.25
            .ForeColor.RGB = RGB(0, 112, 192)  ' 青色
        End With
        
        ' 影の設定
        With .Shadow
            .Visible = msoTrue
            .Style = msoShadowStyleOuterShadow
            .Blur = 5
            .OffsetX = 3
            .OffsetY = 3
            .Transparency = 0.6  ' 60%の透明度
        End With
    End With
End Sub

'******************************
'ユーザー入力テキストを使用するカスタム吹き出し図形を挿入する
'引数  ：なし
'戻り値：なし
'******************************
Sub 吹き出し挿入テキスト入力_EUM()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim userText As String
    
    ' ユーザーからテキスト入力を受け付ける
    userText = InputBox("吹き出しに挿入するテキストを入力してください:", "テキスト入力")
    
    ' キャンセルされた場合は処理を終了
    If userText = "" Then
        MsgBox "キャンセルされました。", vbInformation
        Exit Sub
    End If
    
    ' アクティブシートを設定
    Set ws = ActiveSheet
    
    ' 吹き出し図形を挿入
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangularCallout, 100, 100, 150, 100)
    
    With shp
        ' ユーザー入力テキストを設定
        .TextFrame.characters.Text = userText
        
        ' フォント設定
        With .TextFrame2.TextRange.Font
            .Name = "Arial"
            .Size = 12  ' テキストサイズを調整
            .Fill.ForeColor.RGB = RGB(0, 0, 0)  ' 黒色
        End With
        
        ' テキストの配置
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
        End With
        
        ' 図形の塗りつぶし
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 200)  ' 淡い黄色
            .Transparency = 0.2  ' 20%の透明度
        End With
        
        ' 枠線の設定
        With .Line
            .Visible = msoTrue
            .Weight = 2.25  ' 太さ2.25
            .ForeColor.RGB = RGB(0, 112, 192)  ' 青色
        End With
        
        ' 影の設定
        With .Shadow
            .Visible = msoTrue
            .Style = msoShadowStyleOuterShadow
            .Blur = 5
            .OffsetX = 3
            .OffsetY = 3
            .Transparency = 0.6  ' 60%の透明度
        End With
        
        ' 図形のサイズを自動調整
        .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
    End With
End Sub

'******************************
'赤枠中抜き四角形を挿入する
'引数  ：なし
'戻り値：なし
'******************************
Sub 赤枠中抜き四角形挿入_EUM()
    Dim ws As Worksheet
    Dim shp As Shape
    
    ' アクティブシートを設定
    Set ws = ActiveSheet
    
    ' 四角形を挿入
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 100, 100, 150, 100)
    
    With shp
        ' 塗りつぶしなし
        .Fill.Visible = msoFalse
        
        ' 枠線の設定
        With .Line
            .Visible = msoTrue
            .Weight = 4.5 ' 太さ4.5
            .ForeColor.RGB = RGB(255, 0, 0) ' 赤色
        End With
    End With
End Sub

'******************************
'選択中オブジェクトを赤枠
'引数  ：なし
'戻り値：なし
'******************************
Sub 図形赤枠_EUM()
    Dim shp As Shape
    Dim selectedShapes As ShapeRange
    
    ' 選択中の図形を取得
    On Error Resume Next
    Set selectedShapes = Selection.ShapeRange
    On Error GoTo 0
    
    If Not selectedShapes Is Nothing Then
        ' 選択中のすべての図形に枠線を追加
        For Each shp In selectedShapes
            With shp.Line
                .Visible = msoTrue
                .ForeColor.RGB = RGB(255, 0, 0) ' 赤色
                .Weight = 4 ' 線の太さ：4pt
            End With
        Next shp
    Else
        MsgBox "No shapes are currently selected.", vbExclamation
    End If
End Sub

'******************************
'選択中オブジェクトを枠なし
'引数  ：なし
'戻り値：なし
'******************************
Sub 図形枠なし_EUM()
    Dim shp As Shape
    Dim selectedShapes As ShapeRange
    
    ' 選択中の図形を取得
    On Error Resume Next
    Set selectedShapes = Selection.ShapeRange
    On Error GoTo 0
    
    If Not selectedShapes Is Nothing Then
        ' 選択中のすべての図形の枠線をなしにする
        For Each shp In selectedShapes
            shp.Line.Visible = msoFalse
        Next shp
    Else
        MsgBox "No shapes are currently selected.", vbExclamation
    End If
End Sub

'******************************
'選択中オブジェクトのサイズを50%にする
'引数  ：なし
'戻り値：なし
'******************************
Sub 図形サイズ50_EUM()
    Dim shp As Shape
    Dim selectedShapes As ShapeRange
    
    ' 選択中のオブジェクトを取得
    On Error Resume Next
    Set selectedShapes = Selection.ShapeRange
    On Error GoTo 0
    
    If Not selectedShapes Is Nothing Then
        ' 選択中のすべてのオブジェクトのサイズを50%に縮小
        For Each shp In selectedShapes
            shp.ScaleHeight Factor:=0.5, RelativeToOriginalSize:=msoCTrue
            shp.ScaleWidth Factor:=0.5, RelativeToOriginalSize:=msoCTrue
        Next shp
    Else
        MsgBox "オブジェクトが選択されていません。", vbExclamation
    End If
End Sub


'******************************
'図形と画像を整列する
'引数  ：なし
'戻り値：なし
'******************************
Sub 図形画像整列_EUM()
    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("図形と画像を整列しますか？（警告：戻せません）", vbYesNo + vbQuestion, "確認")
    
    If userResponse = vbYes Then
        On Error GoTo ErrorHandler
    
        ' オブジェクトが選択されているかをチェック
        If Selection Is Nothing Then
            Err.Raise vbObjectError + 1001, , "Please select shapes or pictures before running this macro."
        End If
        
        Dim obj As Object
        Dim topPosition As Double
        Dim leftPosition As Double
        
        topPosition = Selection(1).Top ' 選択したオブジェクトの一番上の位置を取得
        leftPosition = Selection(1).Left ' 選択したオブジェクトの左位置を取得
        
        For Each obj In Selection
            obj.Top = topPosition ' 選択したオブジェクトの一番上の位置に配置
            obj.Left = leftPosition ' 選択したオブジェクトの左位置に配置
            topPosition = topPosition + obj.Height ' オブジェクトの高さを加算して次の位置を計算
        Next obj
        
        Exit Sub
    
    Else
        ShowToast "処理がキャンセルされました。"
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowToast "図形か画像を複数選択して実行しているか確認してください" & vbCrLf & vbCrLf & "An error occurred: " & Err.Description
End Sub


'この関数は、選択されている範囲を定義し返す
'この関数は選択セルだけ
Function GetSelectedRange() As Range
    Dim firstRow As Long, lastRow As Long
    Dim firstColumn As Long, lastColumn As Long
    Dim ws As Worksheet

    Set ws = ActiveSheet
    With Selection
        firstRow = .row
        lastRow = firstRow + .Rows.Count - 1
        firstColumn = .Column
        lastColumn = firstColumn + .Columns.Count - 1
    End With

    Set GetSelectedRange = ws.Range(ws.Cells(firstRow, firstColumn), ws.Cells(lastRow, lastColumn))
End Function
'この関数は、選択セルの最初の行から最後の行までのB:BOの範囲
Function GetSelectedRangeRows() As Range
    Dim firstRow As Long, lastRow As Long
    firstRow = Selection.row
    lastRow = firstRow + Selection.Rows.Count - 1
    Set GetSelectedRangeRows = Range("B" & firstRow & ":BO" & lastRow)
End Function

'******************************
'選択セルを取消線へ
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルを取消線へ_EUM()
    With GetSelectedRange()
        .Font.Strikethrough = True
    End With
End Sub

'******************************
'選択セルをセル色グレーへ
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルをセル色グレーへ_EUM()
    With GetSelectedRange()
        .Interior.Color = RGB(192, 192, 192)  ' グレー
    End With
End Sub

'******************************
'選択セルをセル色グレーと取消線へ
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルをセル色グレーと取消線へ_EUM()
    With GetSelectedRange()
        .Font.Strikethrough = True
        .Interior.Color = RGB(192, 192, 192)  ' グレー
    End With
End Sub

'******************************
'選択セルをセル色レッドへ
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルをセル色レッドへ_EUM()
    With GetSelectedRange()
        .Interior.Color = RGB(255, 0, 0)  ' レッド
    End With
End Sub

'******************************
'選択セルをセル色イエローへ
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルをセル色イエローへ_EUM()
    With GetSelectedRange()
        .Interior.Color = RGB(255, 255, 0)  ' イエロー
    End With
End Sub

'******************************
'選択セルをセル色ブルーへ
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルをセル色ブルーへ_EUM()
    With GetSelectedRange()
        .Interior.Color = RGB(0, 176, 240)  ' ブルー
    End With
End Sub

'******************************
'選択セルをセル色なしへ
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルをセル色なしへ_EUM()
    With GetSelectedRange()
        .Interior.ColorIndex = xlNone
    End With
End Sub

'******************************
'選択セルを取消線なしへ
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルを取消線なしへ_EUM()
    With GetSelectedRange()
        .Font.Strikethrough = False
    End With
End Sub

'******************************
'選択セルをセル色と取消線なしへ
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルをセル色と取消線なしへ_EUM()
    With GetSelectedRange()
        .Font.Strikethrough = False
        .Interior.ColorIndex = xlNone
    End With
End Sub

'******************************
'選択セルの行を取消線へ
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルの行を取消線へ_EUM()
    With GetSelectedRangeRows()
        .Font.Strikethrough = True
    End With
End Sub

'******************************
'選択セルの行をセル色グレーへ
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルの行をセル色グレーへ_EUM()
    With GetSelectedRangeRows()
        .Interior.Color = RGB(192, 192, 192)  ' グレー
    End With
End Sub

'******************************
'選択セルの行をセル色グレーと取消線へ
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルの行をセル色グレーと取消線へ_EUM()
    With GetSelectedRangeRows()
        .Font.Strikethrough = True
        .Interior.Color = RGB(192, 192, 192)  ' グレー
    End With
End Sub

'******************************
'選択セルの行をセル色レッドへ
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルの行をセル色レッドへ_EUM()
    With GetSelectedRangeRows()
        .Interior.Color = RGB(255, 0, 0)  ' レッド
    End With
End Sub

'******************************
'選択セルの行をセル色イエローへ
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルの行をセル色イエローへ_EUM()
    With GetSelectedRangeRows()
        .Interior.Color = RGB(255, 255, 0)  ' イエロー
    End With
End Sub

'******************************
'選択セルの行をセル色ブルーへ
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルの行をセル色ブルーへ_EUM()
    With GetSelectedRangeRows()
        .Interior.Color = RGB(0, 176, 240)  ' ブルー
    End With
End Sub

'******************************
'選択セルの行をセル色なしへ
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルの行をセル色なしへ_EUM()
    With GetSelectedRangeRows()
        .Interior.ColorIndex = xlNone
    End With
End Sub

'******************************
'選択セルの行を取消線なしへ
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルの行を取消線なしへ_EUM()
    With GetSelectedRangeRows()
        .Font.Strikethrough = False
    End With
End Sub

'******************************
'選択セルの行をセル色と取消線なしへ
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルの行をセル色と取消線なしへ_EUM()
    With GetSelectedRangeRows()
        .Font.Strikethrough = False
        .Interior.ColorIndex = xlNone
    End With
End Sub

'******************************
'選択セル行の高さを小さくする
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セル行高小_EUM()
    Dim rng As Range
    Dim row As Range

    ' 現在のアクティブセルまたは選択範囲を取得
    Set rng = Selection

    ' 選択範囲内の各行に対してループ
    For Each row In rng.Rows
        row.RowHeight = row.RowHeight - 8
    Next row
End Sub

'******************************
'選択セル行の高さを大きくする
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セル行高大_EUM()
    Dim rng As Range
    Dim row As Range

    ' 現在のアクティブセルまたは選択範囲を取得
    Set rng = Selection

    ' 選択範囲内の各行に対してループ
    For Each row In rng.Rows
        row.RowHeight = row.RowHeight + 8
    Next row
End Sub


'******************************
'選択セル行コピー挿入
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セル行コピー挿入_EUM()
    Dim selectedRange As Range
    Dim targetRange As Range
    
    On Error Resume Next
    Set selectedRange = Selection.EntireRow
    On Error GoTo 0
    
    If Not selectedRange Is Nothing Then
        Set targetRange = selectedRange.Offset(0).EntireRow
        selectedRange.Copy
        targetRange.Insert
        Application.CutCopyMode = False
    Else
        MsgBox "No row is currently selected.", vbExclamation
    End If
End Sub

'******************************
'選択セル下行追加して行コピー
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セル下行追加して行コピー_EUM()
    Dim selectedRow As Long
    Dim selectedColumn As Long

    ' 選択中の行数をチェック
    If Selection.Rows.Count > 1 Then
        ShowToast "1行のみ選択してください"
        Exit Sub
    End If
    
    ' アクティブセルの行と列を取得
    selectedRow = ActiveCell.row
    selectedColumn = ActiveCell.Column

    ' 選択された行の直下に1行を追加
    Rows(selectedRow + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ' 追加された行全体を選択
    Rows(selectedRow + 1).Select

    ' 上の行の範囲をコピー
    Selection.FillDown

    ' 始めに選択していた行・列を選択
    Cells(selectedRow, selectedColumn).Select
End Sub


'******************************
'選択中セルに連番を振る
'引数  ：なし
'戻り値：なし
'******************************
Sub 連番入力_EUM()
    Dim cell As Range
    Dim counter As Long
    Dim selectedRange As Range
    
    ' 選択中の範囲を取得
    On Error Resume Next
    Set selectedRange = Selection
    On Error GoTo 0
    
    ' 選択範囲が空でないことを確認
    If selectedRange Is Nothing Then
        MsgBox "セルが選択されていません。", vbExclamation
        Exit Sub
    End If
    
    ' カウンターを初期化
    counter = 1
    
    ' 選択された各セルに連番を入力
    For Each cell In selectedRange
        cell.value = counter
        counter = counter + 1
    Next cell
End Sub


'******************************
'選択セルへ連番数式挿入
'選択セル（連番にしたいセル）すぐ上が【数値や日付形式、TRUE（論理値）、エラー値】以外の場合に有効
'引数  ：なし
'戻り値：なし
'******************************
Sub 連番数式挿入1_EUM()
    Selection.FormulaR1C1 = "=IF(ISBLANK(INDIRECT(""RC[1]"",FALSE)), ""-"", N(INDIRECT(""R[-1]C"", FALSE)) + 1)"
End Sub

'******************************
'選択セルへ連番数式挿入
'右セル枝番インクリメント対応
'引数  ：なし
'戻り値：なし
'******************************
Sub 連番数式挿入2_EUM()
    Selection.FormulaR1C1 = "=IF(ISBLANK(INDIRECT(""RC[1]"",FALSE)), ""-"", IF(OR(INDIRECT(""RC[1]"",FALSE)=1, INDIRECT(""RC[1]"",FALSE)=""-""), N(INDIRECT(""R[-1]C"", FALSE)) + 1, N(INDIRECT(""R[-1]C"", FALSE))))"
End Sub


'******************************
'選択セルまでスクロール
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セルまでスクロール_EUM()
    Dim selectedCell As Range
    
    On Error Resume Next
    Set selectedCell = Selection.Cells(1)
    On Error GoTo 0
    
    If Not selectedCell Is Nothing Then
        Dim scrollAmount As Long
        scrollAmount = selectedCell.row
        If scrollAmount >= 5 Then
            scrollAmount = scrollAmount - 5
        Else
            scrollAmount = 1 ' 1行目までスクロール
        End If
        Application.ActiveWindow.ScrollRow = scrollAmount
    Else
        MsgBox "No cell is currently selected.", vbExclamation
    End If
End Sub

'******************************
'すべてのシートを保護
'引数  ：なし
'戻り値：なし
'******************************
Sub シート保護ALL_EUM()
    Dim ws As Worksheet
    
    'すべてのシート
    For Each ws In ActiveWorkbook.Worksheets
        ' シートを保護
        ws.Protect
    Next
    ShowToast "全シート保護しました"
End Sub

'******************************
'選択中のシートを保護
'引数  ：なし
'戻り値：なし
'******************************
Sub シート保護Selected_EUM()
    Dim ws As Worksheet
    Dim strSheetNames As String
    Dim selectedSheetName As String
    
    ' 選択中のシートをループ
    For Each ws In ActiveWindow.SelectedSheets
        ' 選択中のシート名をリストとして保持する
        strSheetNames = strSheetNames & ws.Name & vbCrLf
    Next ws

    ' すべてのシートをループ
    For Each ws In ActiveWorkbook.Worksheets
        ' 選択中のシート名がこのシート名リストに含まれているか確認
        If InStr(1, strSheetNames, ws.Name & vbCrLf) > 0 Then
            ' シートをアクティブにする
            ws.Activate
            ' シートを保護
            ws.Protect
        End If
    Next ws
    ShowToast "選択中シート保護しました"
End Sub

'******************************
'すべてのシートを保護解除
'引数  ：なし
'戻り値：なし
'******************************
Sub シート保護解除ALL_EUM()
    Dim ws As Worksheet
    
    'すべてのシート
    For Each ws In ActiveWorkbook.Worksheets
        ' シートを保護解除
        ws.Unprotect
    Next
    ShowToast "全シート保護解除しました"
End Sub

'******************************
'選択中のシートを保護解除
'引数  ：なし
'戻り値：なし
'******************************
Sub シート保護解除Selected_EUM()
    Dim ws As Worksheet
    Dim strSheetNames As String
    Dim selectedSheetName As String
    
    ' 選択中のシートをループ
    For Each ws In ActiveWindow.SelectedSheets
        ' 選択中のシート名をリストとして保持する
        strSheetNames = strSheetNames & ws.Name & vbCrLf
    Next ws

    ' すべてのシートをループ
    For Each ws In ActiveWorkbook.Worksheets
        ' 選択中のシート名がこのシート名リストに含まれているか確認
        If InStr(1, strSheetNames, ws.Name & vbCrLf) > 0 Then
            ' シートをアクティブにする
            ws.Activate
            ' シートを保護解除
            ws.Unprotect
        End If
    Next ws
    ShowToast "選択中シート保護解除しました"
End Sub

'******************************
'すべての行と列を表示
'引数  ：なし
'戻り値：なし
'******************************
Sub すべて行列表示_EUM()
    ' すべての行と列を表示
    Rows.Hidden = False
    Columns.Hidden = False
    ShowToast "すべての行と列を表示しました"
End Sub

'******************************
'選択中のセルを含む行以外を非表示
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セル行以外非表示_EUM()
    Dim lastRow As Long
    Dim ws As Worksheet

    Set ws = ActiveSheet
    Application.ScreenUpdating = False

    ' データが存在する最後の行を取得
    lastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).row

    ' すべての行を表示
    ws.Rows.Hidden = False

    ' 選択された行以外を非表示
    Dim i As Long
    For i = 1 To lastRow
        If Not Intersect(ws.Rows(i), Selection) Is Nothing Then
            ' 選択されている行はそのままにする
        Else
            ws.Rows(i).Hidden = True
        End If
    Next i

    Application.ScreenUpdating = True
End Sub

'******************************
'選択中のセルを含む列以外を非表示
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択セル列以外非表示_EUM()
    Dim lastCol As Long
    Dim ws As Worksheet

    Set ws = ActiveSheet
    Application.ScreenUpdating = False

    ' データが存在する最後の列を取得
    lastCol = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column

    ' すべての列を表示
    ws.Columns.Hidden = False

    ' 選択された列以外を非表示
    Dim i As Long
    For i = 1 To lastCol
        If Not Intersect(ws.Columns(i), Selection) Is Nothing Then
            ' 選択されている列はそのままにする
        Else
            ws.Columns(i).Hidden = True
        End If
    Next i

    Application.ScreenUpdating = True
End Sub

'******************************
'選択中のセル内の空白でないセル一括選択
'引数  ：なし
'戻り値：なし
'******************************
Sub 空白以外セル一括選択_EUM()
    On Error Resume Next ' エラーが発生した場合にはエラーを無視
    Selection.SpecialCells(xlCellTypeConstants, 23).Select
    On Error GoTo 0 ' エラーハンドリングを元に戻す
End Sub

'******************************
'選択中のセル内の空白セル一括選択
'引数  ：なし
'戻り値：なし
'******************************
Sub 空白セル一括選択_EUM()
    Selection.SpecialCells(xlCellTypeBlanks).Select
End Sub

'******************************
'選択中のセル完全一致セル一括選択
'引数  ：なし
'戻り値：なし
'******************************
Sub 完全一致セル一括選択_EUM()
    Dim matchRange As Range
    Set matchRange = SearchMatchedCells(xlWhole)
    
    ' 見つかったセルがあれば選択
    If Not matchRange Is Nothing Then
        matchRange.Select
    End If
End Sub

'******************************
'選択中のセル部分一致セル一括選択
'引数  ：なし
'戻り値：なし
'******************************
Sub 部分一致セル一括選択_EUM()
    Dim matchRange As Range
    Set matchRange = SearchMatchedCells(xlPart)
    
    ' 見つかったセルがあれば選択
    If Not matchRange Is Nothing Then
        matchRange.Select
    End If
End Sub

' 数式セル選択共通の処理を行うFunction
Function FilterCellsByFormula(rng As Range, includeFormulas As Boolean) As Range
    Dim cell As Range
    Dim resultCells As Range

    For Each cell In rng
        If (cell.HasFormula And includeFormulas) Or (Not cell.HasFormula And Not includeFormulas) Then
            If resultCells Is Nothing Then
                Set resultCells = cell
            Else
                Set resultCells = Union(resultCells, cell)
            End If
        End If
    Next cell

    Set FilterCellsByFormula = resultCells
End Function

'******************************
'選択中のセルから数式セルを除く
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択中セル数式セル除く_EUM()
    Dim nonFormulaCells As Range
    Set nonFormulaCells = FilterCellsByFormula(Selection, False)

    If Not nonFormulaCells Is Nothing Then
        nonFormulaCells.Select
    Else
        ShowToast "選択範囲に非数式セルはありません。"
    End If
End Sub

'******************************
'選択中のセルから数式セルのみ選択する
'引数  ：なし
'戻り値：なし
'******************************
Sub 選択中セル数式セルのみ選択_EUM()
    Dim formulaCells As Range
    Set formulaCells = FilterCellsByFormula(Selection, True)

    If Not formulaCells Is Nothing Then
        formulaCells.Select
    Else
        ShowToast "選択範囲に数式セルはありません。"
    End If
End Sub


'******************************
'指定した値を検索し、一致するセルを返却
'引数  ：
'  lookAtType - 検索方法（xlWholeまたはxlPart）
'戻り値：なし
'******************************
Function SearchMatchedCells(lookAtType As XlLookAt) As Range
    Dim selectedRange As Range
    Dim searchValues As Variant
    Dim foundCell As Range
    Dim combinedRange As Range
    Dim firstAddress As String
    Dim i As Long
    Dim ws As Worksheet
    
    Set ws = ActiveSheet
    Set selectedRange = Selection
    searchValues = CellsValueToArray() ' 選択されたセルの値を配列に格納
    If IsEmpty(searchValues) Then
        Exit Function ' この場合、ここで処理を終了
    End If

    ' アクティブシートの全セルを検索
    For i = LBound(searchValues) To UBound(searchValues)
        Set foundCell = ws.Cells.Find(What:=searchValues(i), After:=ws.Cells(1, 1), LookIn:=xlValues, _
                                      LookAt:=lookAtType, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                      MatchCase:=False)
        If Not foundCell Is Nothing Then
            firstAddress = foundCell.Address
            Do
                ' 検索で見つかったセルを範囲に追加
                If combinedRange Is Nothing Then
                    Set combinedRange = foundCell
                Else
                    Set combinedRange = Union(combinedRange, foundCell)
                End If
                Set foundCell = ws.Cells.FindNext(foundCell)
            Loop While Not foundCell Is Nothing And foundCell.Address <> firstAddress
        End If
    Next i

    ' 見つかったセルがあれば選択し、関数の戻り値として設定
    If Not combinedRange Is Nothing Then
        Set SearchMatchedCells = combinedRange
    Else
        Set SearchMatchedCells = Nothing
    End If
End Function

Sub TestCellsValueToArray()
    Dim myArray As Variant
    Dim i As Long
    Dim toastMessage As String
    
    myArray = CellsValueToArray()
        ' 関数の結果をチェック
    If IsEmpty(myArray) Then
        ' 関数から何も返されなかった場合の処理
        MsgBox "何も返されませんでした。", vbExclamation
        Exit Sub ' この場合、ここで処理を終了
    End If
    ' myArrayを使った処理...
    ' 配列の内容を通知
    toastMessage = ""
    For i = LBound(myArray) To UBound(myArray)
        toastMessage = toastMessage & myArray(i) & " "
    Next i
    
    ' トーストメッセージを表示
    ShowToast toastMessage
End Sub

Function CellsValueToArray() As Variant
    Dim selectedRange As Range
    Dim cell As Range
    Dim resultList As New Collection
    Dim resultArray() As Variant
    Dim i As Long

    ' 選択されているセル範囲を取得
    Set selectedRange = Selection
    
    ' 各セルの値をコレクションに格納
    For Each cell In selectedRange
        If Not IsEmpty(cell.value) Then
            resultList.Add cell.value
        End If
    Next cell
    
    ' resultListが空でないことを確認
    If resultList.Count = 0 Then
        MsgBox "空のセルのみが選択されています。", vbExclamation
        Exit Function ' ここで関数を抜ける
    End If
    
    ' コレクションから配列に変換
    ReDim resultArray(1 To resultList.Count)
    For i = 1 To resultList.Count
        resultArray(i) = resultList(i)
    Next i
    
    ' 配列を戻り値として返却
    CellsValueToArray = resultArray
End Function

'******************************
'選択中のセルの行高さ列幅自動調整
'引数  ：なし
'戻り値：なし
'******************************
Sub 行高さ列幅自動調整_EUM()
    Selection.Rows.AutoFit
    Selection.Columns.AutoFit
End Sub

'******************************
'選択中のセルの値を基に図形を作成
'引数  ：なし
'戻り値：なし
'******************************
Sub 図形作成bySelectedCellValue_EUM()
    Dim cell As Range
    Dim shp As Shape
    Dim currentPrefix As String
    Dim leftPosition As Double
    Dim topPosition As Double
    Dim DataRange As Range
    
    ' 選択中のアクティブセル範囲を指定
    Set DataRange = Selection
    
    ' 初期値として空のプレフィックスを設定
    currentPrefix = ""
    
    ' データのセルをループ
    For Each cell In DataRange
        If cell.Column <> 1 Then
            If cell.Offset(0, -1).value <> "" Then
                ' プレフィックスを更新
                currentPrefix = cell.Offset(0, -1).value & "-"
            End If
        End If
        
        If cell.value <> "" Then
            ' 図形の左上隅の座標を変数化
            leftPosition = cell.Left
            topPosition = cell.Top
            ' 図形を作成しテキストを設定
            ' 図形の作成
            Set shp = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, leftPosition, topPosition, 137.25, 33)
            ' 塗りつぶしの設定
            With shp.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(255, 0, 0)
                .Transparency = 0.8
                .Solid
            End With
            
            ' 枠線の設定
            With shp.Line
                .Visible = msoTrue
                .DashStyle = msoLineSysDash
                .Weight = 1.5
                .ForeColor.RGB = RGB(255, 0, 0)
                .Transparency = 0.5
            End With
            
            ' テキストの設定
            With shp.TextFrame2.TextRange.Font
                .Size = 16
                .NameComplexScript = "HG丸ｺﾞｼｯｸM-PRO"
                .NameFarEast = "HG丸ｺﾞｼｯｸM-PRO"
                .Name = "HG丸ｺﾞｼｯｸM-PRO"
                .Fill.ForeColor.RGB = RGB(255, 0, 0)
                .Fill.Transparency = 0
            End With
            With shp.TextFrame2
                .VerticalAnchor = msoAnchorBottom
                .TextRange.ParagraphFormat.Alignment = msoAlignRight
            End With
            shp.TextFrame.characters.Text = currentPrefix & cell.value
                        
        End If
    Next cell
End Sub
