Attribute VB_Name = "M_MacrosHelper"
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' グローバル変数
Dim GlobalToast As Toast
'******************************
'Toast
'******************************

Sub ShowToast(msg As String)
    Dim i As Double
    Dim startPos As Double
    Dim midPos As Double
    Dim endPos As Double
    Dim duration As Double
    Dim halfway As Double
    Dim currentTime As Double
    Dim changeInPosition As Double

    ' 既存のToastがあればアンロードする
    CloseToast
    
    Set GlobalToast = New Toast
    With GlobalToast
        .messageText = msg
        .StartUpPosition = 0
        .Left = Application.Left + (Application.Width / 2) - (.Width / 2)
        ' トーストの初期位置は画面の下部からさらに10ピクセル下に設定
        startPos = Application.Top + Application.Height - .Height - 50
        ' トーストが最初に移動する中間位置は、初期位置より10ピクセル上に設定
        midPos = startPos - 10
        ' トーストの終了位置は、初期位置からさらに100ピクセル下に設定
        endPos = startPos + 100
        duration = 100 ' アニメーションの総時間を設定
        halfway = duration / 2 ' アニメーションの中間点を設定

        ' UserFormを表示
        .Show vbModeless

        On Error Resume Next ' エラーハンドリングを有効化

        ' アニメーションループ
        For i = 0 To duration
            If Not .Visible Then Exit For ' UserFormが閉じられていたらループを終了
            If Err.Number <> 0 Then Exit For ' 何らかのエラーが発生したらループを終了
            currentTime = i / duration
            If i <= halfway Then
                ' 中間点までの上昇アニメーション
                .Top = startPos - (startPos - midPos) * EaseOut(currentTime * 2)
            Else
                ' 中間点からの下降アニメーション
                .Top = midPos + (endPos - midPos) * EaseIn((currentTime - 0.5) * 2)
            End If
            DoEvents
            Sleep 15
        Next i
        On Error GoTo 0 ' エラーハンドリングを通常モードに戻す
    End With
End Sub

Function EaseOut(t As Double) As Double
    ' イージング関数（EaseOut Cubic）
    t = t - 1
    EaseOut = (t * t * t + 1)
End Function

Function EaseIn(t As Double) As Double
    ' イージング関数（EaseIn Cubic）
    EaseIn = t * t * t
End Function

Sub CloseToast()
    ' すべての開いているToastを閉じる
    Dim frm As Object
    For Each frm In VBA.UserForms
        If TypeName(frm) = "Toast" Then
            Unload frm
        End If
    Next frm
    Set GlobalToast = Nothing
End Sub

