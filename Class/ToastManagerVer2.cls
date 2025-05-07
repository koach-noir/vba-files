Attribute VB_Name = "ToastManagerVer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

' ===================================================
' トースト通知V2 - キュー管理クラス
' ===================================================
' 作成日: 2025/04/25
' 作成者: Claude
' 概要: 改良版トースト通知を管理するためのクラスモジュール
' 機能:
'   - 複数のトースト通知をキュー管理
'   - 通知の種類（情報/成功/警告/エラー）をサポート
'   - 表示時間のカスタマイズ
'   - アニメーション効果の改良
' ===================================================

' トースト通知アイテムを表す型定義
Private Type ToastQueueItem
    Message As String        ' 表示するメッセージ
    ToastType As String      ' 通知の種類（"info", "success", "warning", "error"）
    Duration As Integer      ' 表示時間（ミリ秒）
    IconType As String       ' アイコンタイプ（オプション）
End Type

' プライベート変数
Private toastQueue() As ToastQueueItem   ' トースト通知キュー
Private queueCount As Integer            ' キュー内のアイテム数
Private isProcessing As Boolean          ' 処理中フラグ
Private activeToast As Object            ' 現在表示中のトースト

' 定数
Private Const DEFAULT_DURATION As Integer = 2000    ' デフォルト表示時間（ミリ秒）
Private Const MAX_QUEUE_SIZE As Integer = 10        ' 最大キューサイズ

' ===================================================
' 初期化処理
' ===================================================
Private Sub Class_Initialize()
    ' キューの初期化
    ReDim toastQueue(0 To MAX_QUEUE_SIZE - 1)
    queueCount = 0
    isProcessing = False
End Sub

' ===================================================
' 公開メソッド
' ===================================================

' トースト通知の表示（メイン関数）
Public Sub ShowToast(Message As String, Optional ToastType As String = "info", _
                    Optional Duration As Integer = 0, Optional IconType As String = "")
    ' 引数の検証とデフォルト値の設定
    If Len(Trim(Message)) = 0 Then Exit Sub
    
    If Duration <= 0 Then Duration = DEFAULT_DURATION
    
    ' トーストタイプの検証
    ToastType = ValidateToastType(ToastType)
    
    ' キューに追加
    AddToQueue Message, ToastType, Duration, IconType
    
    ' キュー処理が実行中でなければ開始
    If Not isProcessing Then
        ProcessQueue
    End If
End Sub

' ===================================================
' プライベートメソッド
' ===================================================

' キューに通知を追加
Private Sub AddToQueue(Message As String, ToastType As String, Duration As Integer, IconType As String)
    ' キューがいっぱいなら古いものを削除
    If queueCount >= MAX_QUEUE_SIZE Then
        ' キューを前に詰める
        Dim i As Integer
        For i = 0 To MAX_QUEUE_SIZE - 2
            toastQueue(i) = toastQueue(i + 1)
        Next i
        queueCount = MAX_QUEUE_SIZE - 1
    End If
    
    ' キューに追加
    With toastQueue(queueCount)
        .Message = Message
        .ToastType = ToastType
        .Duration = Duration
        .IconType = IconType
    End With
    
    queueCount = queueCount + 1
End Sub

' キューを処理
Private Sub ProcessQueue()
    If queueCount <= 0 Then
        isProcessing = False
        Exit Sub
    End If
    
    isProcessing = True
    
    ' 最初のトーストを表示
    Dim toast As Object
    Set toast = New ToastVer2
    Set activeToast = toast
    
    With toast
        .Message = toastQueue(0).Message
        .ToastType = toastQueue(0).ToastType
        .Duration = toastQueue(0).Duration
        .IconType = toastQueue(0).IconType
        .Show vbModeless
    End With
    
    ' キューを前に詰める
    Dim i As Integer
    For i = 0 To queueCount - 2
        toastQueue(i) = toastQueue(i + 1)
    Next i
    queueCount = queueCount - 1
    
    ' 完了後に次の処理を行うタイマーをセット
    ' 表示時間 + アニメーション時間(約500ms)を考慮
    Application.OnTime Now + TimeSerial(0, 0, (toastQueue(0).Duration + 500) / 1000), "ContinueToastQueue"
End Sub

' 次のトーストを処理（Application.OnTimeから呼び出される）
Public Sub ContinueToastQueue()
    Set activeToast = Nothing
    ProcessQueue
End Sub

' 現在表示中のトーストを閉じる
Public Sub CloseCurrentToast()
    If Not activeToast Is Nothing Then
        activeToast.Hide
        Set activeToast = Nothing
    End If
    
    ' 次のトーストを処理
    Application.OnTime Now, "ContinueToastQueue"
End Sub

' トーストタイプを検証し、有効な値を返す
Private Function ValidateToastType(ToastType As String) As String
    Select Case LCase(ToastType)
        Case "info", "success", "warning", "error"
            ValidateToastType = LCase(ToastType)
        Case Else
            ValidateToastType = "info"  ' デフォルトは "info"
    End Select
End Function

' ===================================================
' ユーティリティ関数
' ===================================================

' 現在のキュー内アイテム数を取得
Public Property Get QueueSize() As Integer
    QueueSize = queueCount
End Property

' 処理中フラグを取得
Public Property Get IsProcessing() As Boolean
    IsProcessing = isProcessing
End Property

' キューをクリア
Public Sub ClearQueue()
    queueCount = 0
    If Not activeToast Is Nothing Then
        activeToast.Hide
        Set activeToast = Nothing
    End If
    isProcessing = False
End Sub
