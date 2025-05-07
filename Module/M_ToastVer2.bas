Attribute VB_Name = "M_ToastVer2"
Option Explicit

' ===================================================
' トースト通知V2 - ヘルパーモジュール
' ===================================================
' 作成日: 2025/04/25
' 作成者: Claude
' 概要: 改良版トースト通知のヘルパー関数を提供するモジュール
' 機能:
'   - 簡易呼び出し用のヘルパー関数
'   - API関連の宣言
'   - 拡張イージング関数
'   - ユーティリティ関数
' ===================================================

' Windows API宣言
#If VBA7 Then
    ' 64ビット Office用
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    ' 32ビット Office用
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ===================================================
' 公開ヘルパー関数
' ===================================================

' 基本的なトースト通知を表示（簡易バージョン）
Public Sub ShowToastVer2(Message As String, Optional ToastType As String = "info", _
                       Optional Duration As Integer = 0, Optional IconType As String = "")
    ToastManagerVer2.ShowToast Message, ToastType, Duration, IconType
End Sub

' 情報トースト通知（青色）
Public Sub ShowInfoToast(Message As String, Optional Duration As Integer = 0)
    ToastManagerVer2.ShowToast Message, "info", Duration, "info"
End Sub

' 成功トースト通知（緑色）
Public Sub ShowSuccessToast(Message As String, Optional Duration As Integer = 0)
    ToastManagerVer2.ShowToast Message, "success", Duration, "success"
End Sub

' 警告トースト通知（オレンジ色）
Public Sub ShowWarningToast(Message As String, Optional Duration As Integer = 0)
    ToastManagerVer2.ShowToast Message, "warning", Duration, "warning"
End Sub

' エラートースト通知（赤色）
Public Sub ShowErrorToast(Message As String, Optional Duration As Integer = 0)
    ToastManagerVer2.ShowToast Message, "error", Duration, "error"
End Sub

' ===================================================
' 拡張イージング関数
' ===================================================

' 様々なイージング関数を提供（アニメーションの動きをカスタマイズするため）
' t: 0.0～1.0の値（進行度）

' 線形（一定速度）
Public Function LinearEasing(t As Double) As Double
    LinearEasing = t
End Function

' 二次関数イージングイン（加速）
Public Function EaseInQuad(t As Double) As Double
    EaseInQuad = t * t
End Function

' 二次関数イージングアウト（減速）
Public Function EaseOutQuad(t As Double) As Double
    EaseOutQuad = -t * (t - 2)
End Function

' 二次関数イージングインアウト（加速-減速）
Public Function EaseInOutQuad(t As Double) As Double
    t = t * 2
    If t < 1 Then
        EaseInOutQuad = 0.5 * t * t
    Else
        t = t - 1
        EaseInOutQuad = -0.5 * (t * (t - 2) - 1)
    End If
End Function

' 三次関数イージングイン（加速）
Public Function EaseInCubic(t As Double) As Double
    EaseInCubic = t * t * t
End Function

' 三次関数イージングアウト（減速）
Public Function EaseOutCubic(t As Double) As Double
    t = t - 1
    EaseOutCubic = t * t * t + 1
End Function

' 三次関数イージングインアウト（加速-減速）
Public Function EaseInOutCubic(t As Double) As Double
    t = t * 2
    If t < 1 Then
        EaseInOutCubic = 0.5 * t * t * t
    Else
        t = t - 2
        EaseInOutCubic = 0.5 * (t * t * t + 2)
    End If
End Function

' 弾性イージングアウト（バウンド効果）
Public Function EaseOutElastic(t As Double) As Double
    Dim p As Double
    p = 0.3
    
    If t = 0 Then
        EaseOutElastic = 0
        Exit Function
    End If
    
    If t = 1 Then
        EaseOutElastic = 1
        Exit Function
    End If
    
    Dim s As Double
    s = p / 4
    
    EaseOutElastic = 2 ^ (-10 * t) * Sin((t - s) * (2 * WorksheetFunction.Pi) / p) + 1
End Function

' バウンスイージングアウト（跳ね返り効果）
Public Function EaseOutBounce(t As Double) As Double
    If t < (1 / 2.75) Then
        EaseOutBounce = 7.5625 * t * t
    ElseIf t < (2 / 2.75) Then
        t = t - (1.5 / 2.75)
        EaseOutBounce = 7.5625 * t * t + 0.75
    ElseIf t < (2.5 / 2.75) Then
        t = t - (2.25 / 2.75)
        EaseOutBounce = 7.5625 * t * t + 0.9375
    Else
        t = t - (2.625 / 2.75)
        EaseOutBounce = 7.5625 * t * t + 0.984375
    End If
End Function

' ===================================================
' ユーティリティ関数
' ===================================================

' 画面の更新を一時停止/再開する関数
Public Sub FreezeDraw(Optional freeze As Boolean = True)
    If freeze Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    Else
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
    End If
End Sub

' 現在のキュー内トースト数を取得
Public Function GetToastQueueSize() As Integer
    GetToastQueueSize = ToastManagerVer2.QueueSize
End Function

' 現在処理中かどうかを取得
Public Function IsProcessingToasts() As Boolean
    IsProcessingToasts = ToastManagerVer2.IsProcessing
End Function

' トーストキューをクリア
Public Sub ClearToastQueue()
    ToastManagerVer2.ClearQueue
End Sub

' ===================================================
' 使用例
' ===================================================
' 基本的な使用方法:
'  ShowToastVer2 "処理が完了しました", "success", 3000
'
' 専用関数の使用例:
'  ShowSuccessToast "保存しました", 3000
'  ShowErrorToast "エラーが発生しました", 5000
'  ShowInfoToast "情報メッセージ"
'  ShowWarningToast "注意が必要です"
'
' 複数の通知を連続して表示:
'  ShowInfoToast "処理を開始します"
'  ' 何らかの処理
'  ShowSuccessToast "処理が完了しました"
'
' キューのクリア:
'  ClearToastQueue
' ===================================================
