VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ToastVer2 
   Caption         =   "ToastVer2"
   ClientHeight    =   1440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   OleObjectBlob   =   "ToastVer2.frx":0000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   BorderStyle     =   0  'None
End
Attribute VB_Name = "ToastVer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ===================================================
' トースト通知V2 - フォーム
' ===================================================
' 作成日: 2025/04/25
' 作成者: Claude
' 概要: 改良版トースト通知用のユーザーフォーム
' 機能:
'   - 滑らかなアニメーション表示/非表示
'   - 種類別の色分け（情報/成功/警告/エラー）
'   - アイコン表示対応
'   - クリックまたはキー入力で閉じる機能
' ===================================================

' プライベート変数
Private m_Message As String     ' 表示するメッセージ
Private m_ToastType As String   ' 通知の種類（"info", "success", "warning", "error"）
Private m_Duration As Integer   ' 表示時間（ミリ秒）
Private m_IconType As String    ' アイコンタイプ
Private m_AnimationActive As Boolean ' アニメーション進行中フラグ

' アニメーション関連定数
Private Const ANIM_STEP_COUNT As Integer = 20   ' アニメーションのステップ数
Private Const ANIM_STEP_DELAY As Integer = 15   ' ステップ間の遅延（ミリ秒）

' ===================================================
' プロパティ
' ===================================================

' メッセージプロパティ
Public Property Let Message(value As String)
    m_Message = value
    If Me.Visible Then lblMessage.Caption = m_Message
End Property

Public Property Get Message() As String
    Message = m_Message
End Property

' トースト種類プロパティ
Public Property Let ToastType(value As String)
    m_ToastType = value
    ApplyToastStyle
End Property

Public Property Get ToastType() As String
    ToastType = m_ToastType
End Property

' 表示時間プロパティ
Public Property Let Duration(value As Integer)
    m_Duration = value
End Property

Public Property Get Duration() As Integer
    Duration = m_Duration
End Property

' アイコンタイププロパティ
Public Property Let IconType(value As String)
    m_IconType = value
    ApplyIconStyle
End Property

Public Property Get IconType() As String
    IconType = m_IconType
End Property

' ===================================================
' フォームイベント
' ===================================================

' フォーム初期化時
Private Sub UserForm_Initialize()
    ' デフォルト値の設定
    m_ToastType = "info"
    m_Duration = 2000
    m_AnimationActive = False
    
    ' フォームの初期設定
    Me.BackColor = RGB(68, 68, 68)  ' デフォルトの背景色
    Me.Width = 250
    Me.Height = 60
    
    ' 角丸表現のためのフレーム設定
    Frame1.BackColor = Me.BackColor
    Frame1.BorderStyle = 0
    
    ' ラベルの初期設定
    lblMessage.BackColor = Me.BackColor
    lblMessage.ForeColor = RGB(255, 255, 255)
    lblIcon.BackColor = Me.BackColor
    lblIcon.ForeColor = RGB(255, 255, 255)
End Sub

' フォーム表示時
Private Sub UserForm_Activate()
    ' メッセージの設定
    lblMessage.Caption = m_Message
    
    ' スタイルの適用
    ApplyToastStyle
    ApplyIconStyle
    
    ' フォームの位置設定
    PositionForm
    
    ' スライドインアニメーション開始
    SlideInAnimation
    
    ' 表示タイマーの設定
    Application.OnTime Now + TimeSerial(0, 0, m_Duration / 1000), "ToastManagerVer2.ContinueToastQueue"
End Sub

' クリックで閉じる
Private Sub UserForm_Click()
    CloseToast
End Sub

' ラベルクリックでも閉じる
Private Sub lblMessage_Click()
    CloseToast
End Sub

' アイコンラベルクリックでも閉じる
Private Sub lblIcon_Click()
    CloseToast
End Sub

' キー入力でも閉じる
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CloseToast
End Sub

' ===================================================
' メソッド
' ===================================================

' トースト種類に応じたスタイルを適用
Private Sub ApplyToastStyle()
    Dim backColor As Long
    
    ' トースト種類に応じた背景色の設定
    Select Case LCase(m_ToastType)
        Case "info"
            backColor = RGB(0, 120, 215)  ' 青系
        Case "success"
            backColor = RGB(16, 124, 16)  ' 緑系
        Case "warning"
            backColor = RGB(234, 157, 35) ' オレンジ系
        Case "error"
            backColor = RGB(232, 17, 35)  ' 赤系
        Case Else
            backColor = RGB(68, 68, 68)   ' グレー（デフォルト）
    End Select
    
    ' 背景色の適用
    Me.BackColor = backColor
    Frame1.BackColor = backColor
    lblMessage.BackColor = backColor
    lblIcon.BackColor = backColor
End Sub

' アイコン種類に応じたスタイルを適用
Private Sub ApplyIconStyle()
    ' アイコン文字の設定
    ' ここではUnicodeの記号を使用 (PowerPoint風のアイコン)
    Select Case LCase(m_IconType)
        Case "info", ""
            lblIcon.Caption = "i"
            lblIcon.Font.Bold = True
        Case "success"
            lblIcon.Caption = "✓"
        Case "warning"
            lblIcon.Caption = "!"
        Case "error"
            lblIcon.Caption = "×"
        Case Else
            lblIcon.Caption = ""
    End Select
End Sub

' フォームの位置を設定
Private Sub PositionForm()
    ' アクティブウィンドウの中央下部に表示
    ' ExcelのアプリケーションウィンドウではなくアクティブなExcelウィンドウに対して位置決め
    Dim activeWin As Window
    Set activeWin = Application.ActiveWindow
    
    Dim winLeft As Double
    Dim winTop As Double
    Dim winWidth As Double
    Dim winHeight As Double
    
    ' アクティブウィンドウの位置とサイズを取得
    winLeft = activeWin.WindowLeft
    winTop = activeWin.WindowTop
    winWidth = activeWin.Width
    winHeight = activeWin.Height
    
    ' フォームの位置を設定（中央下部）
    Me.Left = winLeft + (winWidth / 2) - (Me.Width / 2)
    Me.Top = winTop + winHeight  ' 画面下（スライドインの開始位置）
End Sub

' スライドインアニメーション
Private Sub SlideInAnimation()
    m_AnimationActive = True
    
    Dim i As Integer
    Dim startTop As Double
    Dim endTop As Double
    Dim currentTop As Double
    
    ' 開始位置と終了位置の設定
    startTop = Me.Top
    endTop = startTop - Me.Height - 10  ' 10ピクセル余裕を持たせる
    
    ' アニメーションループ
    For i = 0 To ANIM_STEP_COUNT
        ' イージング関数を使用した滑らかな動き
        currentTop = startTop - (startTop - endTop) * EaseOutCubic(i / ANIM_STEP_COUNT)
        Me.Top = currentTop
        
        ' 画面の更新と遅延
        DoEvents
        Sleep ANIM_STEP_DELAY
    Next i
    
    m_AnimationActive = False
End Sub

' スライドアウトアニメーション
Private Sub SlideOutAnimation()
    If m_AnimationActive Then Exit Sub
    m_AnimationActive = True
    
    Dim i As Integer
    Dim startTop As Double
    Dim endTop As Double
    Dim currentTop As Double
    
    ' 開始位置と終了位置の設定
    startTop = Me.Top
    endTop = startTop + Me.Height + 10  ' 10ピクセル余裕を持たせる
    
    ' アニメーションループ
    For i = 0 To ANIM_STEP_COUNT
        ' イージング関数を使用した滑らかな動き
        currentTop = startTop + (endTop - startTop) * EaseInCubic(i / ANIM_STEP_COUNT)
        Me.Top = currentTop
        
        ' 画面の更新と遅延
        DoEvents
        Sleep ANIM_STEP_DELAY
    Next i
    
    m_AnimationActive = False
    Me.Hide
End Sub

' トーストを閉じる
Private Sub CloseToast()
    ' アニメーション中なら何もしない
    If m_AnimationActive Then Exit Sub
    
    ' タイマーをキャンセル
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, m_Duration / 1000), "ToastManagerVer2.ContinueToastQueue", , False
    On Error GoTo 0
    
    ' スライドアウトアニメーション
    SlideOutAnimation
    
    ' マネージャーに通知
    ToastManagerVer2.CloseCurrentToast
End Sub

' ===================================================
' イージング関数
' ===================================================

' イージングイン（加速）
Private Function EaseInCubic(t As Double) As Double
    EaseInCubic = t * t * t
End Function

' イージングアウト（減速）
Private Function EaseOutCubic(t As Double) As Double
    t = t - 1
    EaseOutCubic = t * t * t + 1
End Function

' ===================================================
' ユーティリティ関数
' ===================================================

' スリープ関数
Private Sub Sleep(milliseconds As Long)
    Dim startTime As Double
    startTime = Timer
    
    Do While Timer < startTime + (milliseconds / 1000)
        DoEvents
    Loop
End Sub
