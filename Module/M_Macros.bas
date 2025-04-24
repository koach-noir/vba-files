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

    SetClipboard "Hello, ���͐l�Ԃł�"
    
End Sub


'******************************
'�V�[�g�̑I���_�C�A���O��\������
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �V�[�g�I��Dialog_EUM()
   
   With CommandBars.Add(Temporary:=True)
       .Controls.Add(ID:=957).Execute
       .Delete
   End With
   
End Sub

'******************************
'�����o�[�A�w�b�_�[�A�X�N���[���o�[�\���؂�ւ�
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �����w�b�_�X�N���[���\���ؑ�_EUM()

    ' �����o�[�̕\����Ԃ�؂�ւ���
    Application.DisplayFormulaBar = Not Application.DisplayFormulaBar
    With ActiveWindow
        If Application.DisplayFormulaBar Then
            ' �X�N���[���o�[��\��
            .DisplayHorizontalScrollBar = True
            .DisplayVerticalScrollBar = True
            ' �w�b�_�[��\��
            .DisplayHeadings = True
        Else
            ' �X�N���[���o�[���\��
            .DisplayHorizontalScrollBar = False
            .DisplayVerticalScrollBar = False
            ' �w�b�_�[���\��
            .DisplayHeadings = False
        End If
        ShowToast "�\���^��\����؂�ւ��܂���" & vbCrLf & "[scroll][formula][col&row head]"
    End With
End Sub

'******************************
'�S��ʕ\��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �S��ʕ\��_EUM()
    With ActiveWindow
        If Application.DisplayFullScreen Then
        Else
            If Application.DisplayFormulaBar Then
                ' �X�N���[���o�[��\��
                .DisplayHorizontalScrollBar = True
                .DisplayVerticalScrollBar = True
                ' �w�b�_�[��\��
                .DisplayHeadings = True
                ' �S��ʕ\�����[�h�ɐݒ�
                Application.DisplayFullScreen = True
                ' �����o�[�̕\����Ԃ�\��
                Application.DisplayFormulaBar = True
            Else
                ' �X�N���[���o�[���\��
                .DisplayHorizontalScrollBar = False
                .DisplayVerticalScrollBar = False
                ' �w�b�_�[���\��
                .DisplayHeadings = False
                ' �S��ʕ\�����[�h�ɐݒ�
                Application.DisplayFullScreen = True
                ' �����o�[�̕\����Ԃ��\��
                Application.DisplayFormulaBar = False
            End If
        End If
    End With
    ShowToast "�S��ʏI������ɂ�Ecs�L�["
End Sub

'******************************
'�A�N�e�B�u�u�b�N�̖��O���N���b�v�{�[�h�ɃR�s�[
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �u�b�N���R�s�[_EUM()
    Dim bookName As String
    
    ' �A�N�e�B�u�u�b�N�̖��O���擾
    bookName = ActiveWorkbook.Name
    
    ' �N���b�v�{�[�h�Ƀu�b�N�����R�s�[����
    If Len(bookName) > 0 Then
        SetClipboard bookName
        ShowToast "�u�b�N�����N���b�v�{�[�h�ɃR�s�[����܂����B"
    Else
        MsgBox "�u�b�N�����擾�ł��܂���ł����B"
    End If
End Sub

'******************************
'�A�N�e�B�u�u�b�N�̃t���p�X���N���b�v�{�[�h�ɃR�s�[
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �u�b�N�p�X�R�s�[_EUM()
    Dim bookPath As String
    
    ' �A�N�e�B�u�u�b�N�̃t���p�X���擾
    bookPath = ActiveWorkbook.FullName
    
    ' �N���b�v�{�[�h�Ƀu�b�N�̃t���p�X���R�s�[����
    If Len(bookPath) > 0 Then
        SetClipboard bookPath
        ShowToast "�u�b�N�̃t���p�X���N���b�v�{�[�h�ɃR�s�[����܂����B"
    Else
        MsgBox "�u�b�N�̃t���p�X���擾�ł��܂���ł����B"
    End If
End Sub

'******************************
'�A�N�e�B�u�u�b�N�̃p�X���G�N�X�v���[���[�ŕ\��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �u�b�N�p�X���G�N�X�v���[���[�ŕ\��_EUM()
    Dim bookPath As String
    
    ' �A�N�e�B�u�u�b�N�̃t���p�X���擾
    bookPath = ActiveWorkbook.FullName
    
    ' �u�b�N�̃p�X���擾�ł����ꍇ�̓G�N�X�v���[���[�ŕ\������
    If Len(bookPath) > 0 Then
        Shell "explorer.exe /select," & bookPath, vbNormalFocus
    Else
        MsgBox "�u�b�N�̃p�X���擾�ł��܂���ł����B"
    End If
End Sub


'******************************
'�I���Z���̓��e�̖��O�ŃV�[�g�쐬
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �V�[�g�쐬SelectCell2SheetName_EUM()
    Dim baseSheetName As String
    Dim sheetName As String
    Dim sheetNumber As Integer
    Dim newSheet As Worksheet
    Dim originalSheet As Worksheet
    Dim cell As Range

    ' ���݂̃A�N�e�B�u�V�[�g��ۑ�
    Set originalSheet = ActiveSheet

    ' �I��͈͓��̊e�Z���ɑ΂��ď���
    For Each cell In Selection
        baseSheetName = ReplaceInvalidCharacters(cell.Text)

        ' �Œ�31�����ɐ���
        If Len(baseSheetName) > 31 Then
            baseSheetName = Left(baseSheetName, 31)
        End If

        sheetNumber = 1
        sheetName = baseSheetName

        ' �V�[�g�������ɑ��݂���ꍇ�͔ԍ��𑝂₷
        While SheetExistsInActiveWorkbook(sheetName)
            sheetNumber = sheetNumber + 1
            sheetName = baseSheetName & "(" & sheetNumber & ")"
        Wend

        ' �V�����V�[�g���쐬
        Set newSheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        
        If sheetName <> "" Then newSheet.Name = sheetName
        
        ' A1�ɃV�[�g��
        newSheet.Range("A1").Formula = "=MID(@CELL(""filename"",A1),FIND(""]"",@CELL(""filename"",A1))+1,31)"

        ' A2�ɐ�����ݒ�
        newSheet.Range("A2").Formula = "=IFERROR(HYPERLINK(""#'" & """ & B2 & ""'!A2"", "" << ""), """")"
        With newSheet.Range("A2").Font
            .Name = "���C���I"  ' �t�H���g�����C���I�ɐݒ�
            .Color = RGB(0, 112, 192)  ' �����F��HEX #0070C0�iRGB��0, 112, 192�j�ɐݒ�
            .Underline = xlUnderlineStyleSingle  ' ������ݒ�
        End With

        ' B2�Ɍ��̃V�[�g����ݒ�
        newSheet.Range("B2").value = originalSheet.Name
        
    Next cell

    ' ���̃V�[�g�ɖ߂�
    originalSheet.Activate
End Sub

' �V�[�g���Ɏg�p�ł��Ȃ����������O����֐�
Function ReplaceInvalidCharacters(inputString As String) As String
    Dim invalidChars As Variant
    invalidChars = Array("/", "\", ":", "*", "?", """", "<", ">", "|", "�E", Chr(10), Chr(13))

    Dim i As Integer
    For i = LBound(invalidChars) To UBound(invalidChars)
        inputString = Replace(inputString, invalidChars(i), "")
    Next i

    ReplaceInvalidCharacters = inputString
End Function

' �A�N�e�B�u�ȃ��[�N�u�b�N���ŃV�[�g�����݂��邩�ǂ������m�F����w���p�[�֐�
Function SheetExistsInActiveWorkbook(sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ActiveWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExistsInActiveWorkbook = Not sheet Is Nothing
End Function

'******************************
'�S�ẴV�[�g��A1�I��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �S�ẴV�[�g��A1�I��_EUM()
    Dim ws As Worksheet
    Dim firstSheet As Worksheet

    ' �u�b�N���̐擪�V�[�g���擾
    Set firstSheet = ActiveWorkbook.Sheets(1)

    ' �A�v���P�[�V�����̃X�N���[���A�b�v�f�[�g���~����i�������������j
    Application.ScreenUpdating = False

    ' ���ׂẴV�[�g�����[�v�ŏ���
    For Each ws In ActiveWorkbook.Sheets
        ' �V�[�g���A�N�e�B�u�ɂ���A1�Z�����A�N�e�B�u�ɂ���
        ws.Activate
        ws.Range("A1").Activate
        
        ' �\���{����80%�ɐݒ�
        ActiveWindow.Zoom = 80
        
        ' �\���������[�փX�N���[��
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    Next ws

    ' �A�v���P�[�V�����̃X�N���[���A�b�v�f�[�g���ĊJ����
    Application.ScreenUpdating = True

    ' �u�b�N���̐擪�V�[�g���A�N�e�B�u�ɂ��ĕ\��
    firstSheet.Activate

    ShowToast "�S�V�[�gA1�Z���I���A�\���{���F80%"
End Sub


'******************************
'���ׂẴV�[�g�����N���b�v�{�[�h�ɃR�s�[
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �V�[�g���R�s�[ALL_EUM()
    Dim ws As Worksheet
    Dim strSheetNames As String
    
    '���ׂẴV�[�g����A��������������쐬����
    For Each ws In ActiveWorkbook.Worksheets
        strSheetNames = strSheetNames & ws.Name & vbCrLf
    Next
    
    ' �N���b�v�{�[�h�ɃV�[�g�����R�s�[����
    If Len(strSheetNames) > 0 Then
        SetClipboard strSheetNames
        ShowToast "���ׂẴV�[�g�����N���b�v�{�[�h�ɃR�s�[����܂����B"
    Else
        MsgBox "���̃u�b�N�ɂ̓V�[�g������܂���B"
    End If
End Sub

'******************************
'�I�𒆂̃V�[�g�����N���b�v�{�[�h�ɃR�s�[
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �V�[�g���R�s�[Selected_EUM()
    Dim ws As Worksheet
    Dim strSheetNames As String

    ' �I�𒆂̃V�[�g����A��������������쐬����
    For Each ws In ActiveWindow.SelectedSheets
        strSheetNames = strSheetNames & ws.Name & vbCrLf
    Next

    ' �N���b�v�{�[�h�ɃV�[�g�����R�s�[����
    If Len(strSheetNames) > 0 Then
        SetClipboard strSheetNames
        ShowToast "�I�����ꂽ�V�[�g�����N���b�v�{�[�h�ɃR�s�[����܂����B"
    Else
        MsgBox "�I�����ꂽ�V�[�g������܂���B"
    End If
End Sub

'******************************
'�I�𒆃I�u�W�F�N�g���Ŕw�ʂ�
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I�u�W�F�N�g���Ŕw�ʂ�_EUM()
    On Error GoTo ErrorHandler
    Selection.ShapeRange.ZOrder msoSendToBack
    Exit Sub
ErrorHandler:
    Application.Wait Now + TimeValue("0:00:01") ' 1�b�ҋ@
End Sub

'******************************
'�I�𒆃I�u�W�F�N�g���őO�ʂ�
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I�u�W�F�N�g���őO�ʂ�_EUM()
    On Error GoTo ErrorHandler
    Selection.ShapeRange.ZOrder msoBringToFront
    Exit Sub
ErrorHandler:
    Application.Wait Now + TimeValue("0:00:01") ' 1�b�ҋ@
End Sub

'******************************
'�����o���}�`��}������
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �����o���}��_EUM()
    Dim ws As Worksheet
    Dim shp As Shape
    
    ' �A�N�e�B�u�V�[�g��ݒ�
    Set ws = ActiveSheet
    
    ' �����o���}�`��}��
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangularCallout, 100, 100, 150, 100)
    
    With shp
        ' �e�L�X�g��ݒ�
        .TextFrame.characters.Text = "*"
        
        ' �t�H���g�ݒ�
        With .TextFrame2.TextRange.Font
            .Name = "Arial"
            .Size = 24
            .Fill.ForeColor.RGB = RGB(0, 0, 0)  ' ���F
        End With
        
        ' �}�`�̓h��Ԃ�
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 200)  ' �W�����F
            .Transparency = 0.2  ' 20%�̓����x
        End With
        
        ' �g���̐ݒ�
        With .Line
            .Visible = msoTrue
            .Weight = 2.25  ' ����2.25
            .ForeColor.RGB = RGB(0, 112, 192)  ' �F
        End With
        
        ' �e�̐ݒ�
        With .Shadow
            .Visible = msoTrue
            .Style = msoShadowStyleOuterShadow
            .Blur = 5
            .OffsetX = 3
            .OffsetY = 3
            .Transparency = 0.6  ' 60%�̓����x
        End With
    End With
End Sub

'******************************
'���[�U�[���̓e�L�X�g���g�p����J�X�^�������o���}�`��}������
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �����o���}���e�L�X�g����_EUM()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim userText As String
    
    ' ���[�U�[����e�L�X�g���͂��󂯕t����
    userText = InputBox("�����o���ɑ}������e�L�X�g����͂��Ă�������:", "�e�L�X�g����")
    
    ' �L�����Z�����ꂽ�ꍇ�͏������I��
    If userText = "" Then
        MsgBox "�L�����Z������܂����B", vbInformation
        Exit Sub
    End If
    
    ' �A�N�e�B�u�V�[�g��ݒ�
    Set ws = ActiveSheet
    
    ' �����o���}�`��}��
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangularCallout, 100, 100, 150, 100)
    
    With shp
        ' ���[�U�[���̓e�L�X�g��ݒ�
        .TextFrame.characters.Text = userText
        
        ' �t�H���g�ݒ�
        With .TextFrame2.TextRange.Font
            .Name = "Arial"
            .Size = 12  ' �e�L�X�g�T�C�Y�𒲐�
            .Fill.ForeColor.RGB = RGB(0, 0, 0)  ' ���F
        End With
        
        ' �e�L�X�g�̔z�u
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
        End With
        
        ' �}�`�̓h��Ԃ�
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 200)  ' �W�����F
            .Transparency = 0.2  ' 20%�̓����x
        End With
        
        ' �g���̐ݒ�
        With .Line
            .Visible = msoTrue
            .Weight = 2.25  ' ����2.25
            .ForeColor.RGB = RGB(0, 112, 192)  ' �F
        End With
        
        ' �e�̐ݒ�
        With .Shadow
            .Visible = msoTrue
            .Style = msoShadowStyleOuterShadow
            .Blur = 5
            .OffsetX = 3
            .OffsetY = 3
            .Transparency = 0.6  ' 60%�̓����x
        End With
        
        ' �}�`�̃T�C�Y����������
        .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
    End With
End Sub

'******************************
'�Ԙg�������l�p�`��}������
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �Ԙg�������l�p�`�}��_EUM()
    Dim ws As Worksheet
    Dim shp As Shape
    
    ' �A�N�e�B�u�V�[�g��ݒ�
    Set ws = ActiveSheet
    
    ' �l�p�`��}��
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 100, 100, 150, 100)
    
    With shp
        ' �h��Ԃ��Ȃ�
        .Fill.Visible = msoFalse
        
        ' �g���̐ݒ�
        With .Line
            .Visible = msoTrue
            .Weight = 4.5 ' ����4.5
            .ForeColor.RGB = RGB(255, 0, 0) ' �ԐF
        End With
    End With
End Sub

'******************************
'�I�𒆃I�u�W�F�N�g��Ԙg
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �}�`�Ԙg_EUM()
    Dim shp As Shape
    Dim selectedShapes As ShapeRange
    
    ' �I�𒆂̐}�`���擾
    On Error Resume Next
    Set selectedShapes = Selection.ShapeRange
    On Error GoTo 0
    
    If Not selectedShapes Is Nothing Then
        ' �I�𒆂̂��ׂĂ̐}�`�ɘg����ǉ�
        For Each shp In selectedShapes
            With shp.Line
                .Visible = msoTrue
                .ForeColor.RGB = RGB(255, 0, 0) ' �ԐF
                .Weight = 4 ' ���̑����F4pt
            End With
        Next shp
    Else
        MsgBox "No shapes are currently selected.", vbExclamation
    End If
End Sub

'******************************
'�I�𒆃I�u�W�F�N�g��g�Ȃ�
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �}�`�g�Ȃ�_EUM()
    Dim shp As Shape
    Dim selectedShapes As ShapeRange
    
    ' �I�𒆂̐}�`���擾
    On Error Resume Next
    Set selectedShapes = Selection.ShapeRange
    On Error GoTo 0
    
    If Not selectedShapes Is Nothing Then
        ' �I�𒆂̂��ׂĂ̐}�`�̘g�����Ȃ��ɂ���
        For Each shp In selectedShapes
            shp.Line.Visible = msoFalse
        Next shp
    Else
        MsgBox "No shapes are currently selected.", vbExclamation
    End If
End Sub

'******************************
'�I�𒆃I�u�W�F�N�g�̃T�C�Y��50%�ɂ���
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �}�`�T�C�Y50_EUM()
    Dim shp As Shape
    Dim selectedShapes As ShapeRange
    
    ' �I�𒆂̃I�u�W�F�N�g���擾
    On Error Resume Next
    Set selectedShapes = Selection.ShapeRange
    On Error GoTo 0
    
    If Not selectedShapes Is Nothing Then
        ' �I�𒆂̂��ׂẴI�u�W�F�N�g�̃T�C�Y��50%�ɏk��
        For Each shp In selectedShapes
            shp.ScaleHeight Factor:=0.5, RelativeToOriginalSize:=msoCTrue
            shp.ScaleWidth Factor:=0.5, RelativeToOriginalSize:=msoCTrue
        Next shp
    Else
        MsgBox "�I�u�W�F�N�g���I������Ă��܂���B", vbExclamation
    End If
End Sub


'******************************
'�}�`�Ɖ摜�𐮗񂷂�
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �}�`�摜����_EUM()
    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("�}�`�Ɖ摜�𐮗񂵂܂����H�i�x���F�߂��܂���j", vbYesNo + vbQuestion, "�m�F")
    
    If userResponse = vbYes Then
        On Error GoTo ErrorHandler
    
        ' �I�u�W�F�N�g���I������Ă��邩���`�F�b�N
        If Selection Is Nothing Then
            Err.Raise vbObjectError + 1001, , "Please select shapes or pictures before running this macro."
        End If
        
        Dim obj As Object
        Dim topPosition As Double
        Dim leftPosition As Double
        
        topPosition = Selection(1).Top ' �I�������I�u�W�F�N�g�̈�ԏ�̈ʒu���擾
        leftPosition = Selection(1).Left ' �I�������I�u�W�F�N�g�̍��ʒu���擾
        
        For Each obj In Selection
            obj.Top = topPosition ' �I�������I�u�W�F�N�g�̈�ԏ�̈ʒu�ɔz�u
            obj.Left = leftPosition ' �I�������I�u�W�F�N�g�̍��ʒu�ɔz�u
            topPosition = topPosition + obj.Height ' �I�u�W�F�N�g�̍��������Z���Ď��̈ʒu���v�Z
        Next obj
        
        Exit Sub
    
    Else
        ShowToast "�������L�����Z������܂����B"
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowToast "�}�`���摜�𕡐��I�����Ď��s���Ă��邩�m�F���Ă�������" & vbCrLf & vbCrLf & "An error occurred: " & Err.Description
End Sub


'���̊֐��́A�I������Ă���͈͂��`���Ԃ�
'���̊֐��͑I���Z������
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
'���̊֐��́A�I���Z���̍ŏ��̍s����Ō�̍s�܂ł�B:BO�͈̔�
Function GetSelectedRangeRows() As Range
    Dim firstRow As Long, lastRow As Long
    firstRow = Selection.row
    lastRow = firstRow + Selection.Rows.Count - 1
    Set GetSelectedRangeRows = Range("B" & firstRow & ":BO" & lastRow)
End Function

'******************************
'�I���Z�����������
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z�����������_EUM()
    With GetSelectedRange()
        .Font.Strikethrough = True
    End With
End Sub

'******************************
'�I���Z�����Z���F�O���[��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z�����Z���F�O���[��_EUM()
    With GetSelectedRange()
        .Interior.Color = RGB(192, 192, 192)  ' �O���[
    End With
End Sub

'******************************
'�I���Z�����Z���F�O���[�Ǝ������
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z�����Z���F�O���[�Ǝ������_EUM()
    With GetSelectedRange()
        .Font.Strikethrough = True
        .Interior.Color = RGB(192, 192, 192)  ' �O���[
    End With
End Sub

'******************************
'�I���Z�����Z���F���b�h��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z�����Z���F���b�h��_EUM()
    With GetSelectedRange()
        .Interior.Color = RGB(255, 0, 0)  ' ���b�h
    End With
End Sub

'******************************
'�I���Z�����Z���F�C�G���[��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z�����Z���F�C�G���[��_EUM()
    With GetSelectedRange()
        .Interior.Color = RGB(255, 255, 0)  ' �C�G���[
    End With
End Sub

'******************************
'�I���Z�����Z���F�u���[��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z�����Z���F�u���[��_EUM()
    With GetSelectedRange()
        .Interior.Color = RGB(0, 176, 240)  ' �u���[
    End With
End Sub

'******************************
'�I���Z�����Z���F�Ȃ���
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z�����Z���F�Ȃ���_EUM()
    With GetSelectedRange()
        .Interior.ColorIndex = xlNone
    End With
End Sub

'******************************
'�I���Z����������Ȃ���
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z����������Ȃ���_EUM()
    With GetSelectedRange()
        .Font.Strikethrough = False
    End With
End Sub

'******************************
'�I���Z�����Z���F�Ǝ�����Ȃ���
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z�����Z���F�Ǝ�����Ȃ���_EUM()
    With GetSelectedRange()
        .Font.Strikethrough = False
        .Interior.ColorIndex = xlNone
    End With
End Sub

'******************************
'�I���Z���̍s���������
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z���̍s���������_EUM()
    With GetSelectedRangeRows()
        .Font.Strikethrough = True
    End With
End Sub

'******************************
'�I���Z���̍s���Z���F�O���[��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z���̍s���Z���F�O���[��_EUM()
    With GetSelectedRangeRows()
        .Interior.Color = RGB(192, 192, 192)  ' �O���[
    End With
End Sub

'******************************
'�I���Z���̍s���Z���F�O���[�Ǝ������
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z���̍s���Z���F�O���[�Ǝ������_EUM()
    With GetSelectedRangeRows()
        .Font.Strikethrough = True
        .Interior.Color = RGB(192, 192, 192)  ' �O���[
    End With
End Sub

'******************************
'�I���Z���̍s���Z���F���b�h��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z���̍s���Z���F���b�h��_EUM()
    With GetSelectedRangeRows()
        .Interior.Color = RGB(255, 0, 0)  ' ���b�h
    End With
End Sub

'******************************
'�I���Z���̍s���Z���F�C�G���[��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z���̍s���Z���F�C�G���[��_EUM()
    With GetSelectedRangeRows()
        .Interior.Color = RGB(255, 255, 0)  ' �C�G���[
    End With
End Sub

'******************************
'�I���Z���̍s���Z���F�u���[��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z���̍s���Z���F�u���[��_EUM()
    With GetSelectedRangeRows()
        .Interior.Color = RGB(0, 176, 240)  ' �u���[
    End With
End Sub

'******************************
'�I���Z���̍s���Z���F�Ȃ���
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z���̍s���Z���F�Ȃ���_EUM()
    With GetSelectedRangeRows()
        .Interior.ColorIndex = xlNone
    End With
End Sub

'******************************
'�I���Z���̍s��������Ȃ���
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z���̍s��������Ȃ���_EUM()
    With GetSelectedRangeRows()
        .Font.Strikethrough = False
    End With
End Sub

'******************************
'�I���Z���̍s���Z���F�Ǝ�����Ȃ���
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z���̍s���Z���F�Ǝ�����Ȃ���_EUM()
    With GetSelectedRangeRows()
        .Font.Strikethrough = False
        .Interior.ColorIndex = xlNone
    End With
End Sub

'******************************
'�I���Z���s�̍���������������
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z���s����_EUM()
    Dim rng As Range
    Dim row As Range

    ' ���݂̃A�N�e�B�u�Z���܂��͑I��͈͂��擾
    Set rng = Selection

    ' �I��͈͓��̊e�s�ɑ΂��ă��[�v
    For Each row In rng.Rows
        row.RowHeight = row.RowHeight - 8
    Next row
End Sub

'******************************
'�I���Z���s�̍�����傫������
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z���s����_EUM()
    Dim rng As Range
    Dim row As Range

    ' ���݂̃A�N�e�B�u�Z���܂��͑I��͈͂��擾
    Set rng = Selection

    ' �I��͈͓��̊e�s�ɑ΂��ă��[�v
    For Each row In rng.Rows
        row.RowHeight = row.RowHeight + 8
    Next row
End Sub


'******************************
'�I���Z���s�R�s�[�}��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z���s�R�s�[�}��_EUM()
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
'�I���Z�����s�ǉ����čs�R�s�[
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z�����s�ǉ����čs�R�s�[_EUM()
    Dim selectedRow As Long
    Dim selectedColumn As Long

    ' �I�𒆂̍s�����`�F�b�N
    If Selection.Rows.Count > 1 Then
        ShowToast "1�s�̂ݑI�����Ă�������"
        Exit Sub
    End If
    
    ' �A�N�e�B�u�Z���̍s�Ɨ���擾
    selectedRow = ActiveCell.row
    selectedColumn = ActiveCell.Column

    ' �I�����ꂽ�s�̒�����1�s��ǉ�
    Rows(selectedRow + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ' �ǉ����ꂽ�s�S�̂�I��
    Rows(selectedRow + 1).Select

    ' ��̍s�͈̔͂��R�s�[
    Selection.FillDown

    ' �n�߂ɑI�����Ă����s�E���I��
    Cells(selectedRow, selectedColumn).Select
End Sub


'******************************
'�I�𒆃Z���ɘA�Ԃ�U��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �A�ԓ���_EUM()
    Dim cell As Range
    Dim counter As Long
    Dim selectedRange As Range
    
    ' �I�𒆂͈̔͂��擾
    On Error Resume Next
    Set selectedRange = Selection
    On Error GoTo 0
    
    ' �I��͈͂���łȂ����Ƃ��m�F
    If selectedRange Is Nothing Then
        MsgBox "�Z�����I������Ă��܂���B", vbExclamation
        Exit Sub
    End If
    
    ' �J�E���^�[��������
    counter = 1
    
    ' �I�����ꂽ�e�Z���ɘA�Ԃ����
    For Each cell In selectedRange
        cell.value = counter
        counter = counter + 1
    Next cell
End Sub


'******************************
'�I���Z���֘A�Ԑ����}��
'�I���Z���i�A�Ԃɂ������Z���j�����オ�y���l����t�`���ATRUE�i�_���l�j�A�G���[�l�z�ȊO�̏ꍇ�ɗL��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �A�Ԑ����}��1_EUM()
    Selection.FormulaR1C1 = "=IF(ISBLANK(INDIRECT(""RC[1]"",FALSE)), ""-"", N(INDIRECT(""R[-1]C"", FALSE)) + 1)"
End Sub

'******************************
'�I���Z���֘A�Ԑ����}��
'�E�Z���}�ԃC���N�������g�Ή�
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �A�Ԑ����}��2_EUM()
    Selection.FormulaR1C1 = "=IF(ISBLANK(INDIRECT(""RC[1]"",FALSE)), ""-"", IF(OR(INDIRECT(""RC[1]"",FALSE)=1, INDIRECT(""RC[1]"",FALSE)=""-""), N(INDIRECT(""R[-1]C"", FALSE)) + 1, N(INDIRECT(""R[-1]C"", FALSE))))"
End Sub


'******************************
'�I���Z���܂ŃX�N���[��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z���܂ŃX�N���[��_EUM()
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
            scrollAmount = 1 ' 1�s�ڂ܂ŃX�N���[��
        End If
        Application.ActiveWindow.ScrollRow = scrollAmount
    Else
        MsgBox "No cell is currently selected.", vbExclamation
    End If
End Sub

'******************************
'���ׂẴV�[�g��ی�
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �V�[�g�ی�ALL_EUM()
    Dim ws As Worksheet
    
    '���ׂẴV�[�g
    For Each ws In ActiveWorkbook.Worksheets
        ' �V�[�g��ی�
        ws.Protect
    Next
    ShowToast "�S�V�[�g�ی삵�܂���"
End Sub

'******************************
'�I�𒆂̃V�[�g��ی�
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �V�[�g�ی�Selected_EUM()
    Dim ws As Worksheet
    Dim strSheetNames As String
    Dim selectedSheetName As String
    
    ' �I�𒆂̃V�[�g�����[�v
    For Each ws In ActiveWindow.SelectedSheets
        ' �I�𒆂̃V�[�g�������X�g�Ƃ��ĕێ�����
        strSheetNames = strSheetNames & ws.Name & vbCrLf
    Next ws

    ' ���ׂẴV�[�g�����[�v
    For Each ws In ActiveWorkbook.Worksheets
        ' �I�𒆂̃V�[�g�������̃V�[�g�����X�g�Ɋ܂܂�Ă��邩�m�F
        If InStr(1, strSheetNames, ws.Name & vbCrLf) > 0 Then
            ' �V�[�g���A�N�e�B�u�ɂ���
            ws.Activate
            ' �V�[�g��ی�
            ws.Protect
        End If
    Next ws
    ShowToast "�I�𒆃V�[�g�ی삵�܂���"
End Sub

'******************************
'���ׂẴV�[�g��ی����
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �V�[�g�ی����ALL_EUM()
    Dim ws As Worksheet
    
    '���ׂẴV�[�g
    For Each ws In ActiveWorkbook.Worksheets
        ' �V�[�g��ی����
        ws.Unprotect
    Next
    ShowToast "�S�V�[�g�ی�������܂���"
End Sub

'******************************
'�I�𒆂̃V�[�g��ی����
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �V�[�g�ی����Selected_EUM()
    Dim ws As Worksheet
    Dim strSheetNames As String
    Dim selectedSheetName As String
    
    ' �I�𒆂̃V�[�g�����[�v
    For Each ws In ActiveWindow.SelectedSheets
        ' �I�𒆂̃V�[�g�������X�g�Ƃ��ĕێ�����
        strSheetNames = strSheetNames & ws.Name & vbCrLf
    Next ws

    ' ���ׂẴV�[�g�����[�v
    For Each ws In ActiveWorkbook.Worksheets
        ' �I�𒆂̃V�[�g�������̃V�[�g�����X�g�Ɋ܂܂�Ă��邩�m�F
        If InStr(1, strSheetNames, ws.Name & vbCrLf) > 0 Then
            ' �V�[�g���A�N�e�B�u�ɂ���
            ws.Activate
            ' �V�[�g��ی����
            ws.Unprotect
        End If
    Next ws
    ShowToast "�I�𒆃V�[�g�ی�������܂���"
End Sub

'******************************
'���ׂĂ̍s�Ɨ��\��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub ���ׂčs��\��_EUM()
    ' ���ׂĂ̍s�Ɨ��\��
    Rows.Hidden = False
    Columns.Hidden = False
    ShowToast "���ׂĂ̍s�Ɨ��\�����܂���"
End Sub

'******************************
'�I�𒆂̃Z�����܂ލs�ȊO���\��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z���s�ȊO��\��_EUM()
    Dim lastRow As Long
    Dim ws As Worksheet

    Set ws = ActiveSheet
    Application.ScreenUpdating = False

    ' �f�[�^�����݂���Ō�̍s���擾
    lastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).row

    ' ���ׂĂ̍s��\��
    ws.Rows.Hidden = False

    ' �I�����ꂽ�s�ȊO���\��
    Dim i As Long
    For i = 1 To lastRow
        If Not Intersect(ws.Rows(i), Selection) Is Nothing Then
            ' �I������Ă���s�͂��̂܂܂ɂ���
        Else
            ws.Rows(i).Hidden = True
        End If
    Next i

    Application.ScreenUpdating = True
End Sub

'******************************
'�I�𒆂̃Z�����܂ޗ�ȊO���\��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I���Z����ȊO��\��_EUM()
    Dim lastCol As Long
    Dim ws As Worksheet

    Set ws = ActiveSheet
    Application.ScreenUpdating = False

    ' �f�[�^�����݂���Ō�̗���擾
    lastCol = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column

    ' ���ׂĂ̗��\��
    ws.Columns.Hidden = False

    ' �I�����ꂽ��ȊO���\��
    Dim i As Long
    For i = 1 To lastCol
        If Not Intersect(ws.Columns(i), Selection) Is Nothing Then
            ' �I������Ă����͂��̂܂܂ɂ���
        Else
            ws.Columns(i).Hidden = True
        End If
    Next i

    Application.ScreenUpdating = True
End Sub

'******************************
'�I�𒆂̃Z�����̋󔒂łȂ��Z���ꊇ�I��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �󔒈ȊO�Z���ꊇ�I��_EUM()
    On Error Resume Next ' �G���[�����������ꍇ�ɂ̓G���[�𖳎�
    Selection.SpecialCells(xlCellTypeConstants, 23).Select
    On Error GoTo 0 ' �G���[�n���h�����O�����ɖ߂�
End Sub

'******************************
'�I�𒆂̃Z�����̋󔒃Z���ꊇ�I��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �󔒃Z���ꊇ�I��_EUM()
    Selection.SpecialCells(xlCellTypeBlanks).Select
End Sub

'******************************
'�I�𒆂̃Z�����S��v�Z���ꊇ�I��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub ���S��v�Z���ꊇ�I��_EUM()
    Dim matchRange As Range
    Set matchRange = SearchMatchedCells(xlWhole)
    
    ' ���������Z��������ΑI��
    If Not matchRange Is Nothing Then
        matchRange.Select
    End If
End Sub

'******************************
'�I�𒆂̃Z��������v�Z���ꊇ�I��
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub ������v�Z���ꊇ�I��_EUM()
    Dim matchRange As Range
    Set matchRange = SearchMatchedCells(xlPart)
    
    ' ���������Z��������ΑI��
    If Not matchRange Is Nothing Then
        matchRange.Select
    End If
End Sub

' �����Z���I�����ʂ̏������s��Function
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
'�I�𒆂̃Z�����琔���Z��������
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I�𒆃Z�������Z������_EUM()
    Dim nonFormulaCells As Range
    Set nonFormulaCells = FilterCellsByFormula(Selection, False)

    If Not nonFormulaCells Is Nothing Then
        nonFormulaCells.Select
    Else
        ShowToast "�I��͈͂ɔ񐔎��Z���͂���܂���B"
    End If
End Sub

'******************************
'�I�𒆂̃Z�����琔���Z���̂ݑI������
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �I�𒆃Z�������Z���̂ݑI��_EUM()
    Dim formulaCells As Range
    Set formulaCells = FilterCellsByFormula(Selection, True)

    If Not formulaCells Is Nothing Then
        formulaCells.Select
    Else
        ShowToast "�I��͈͂ɐ����Z���͂���܂���B"
    End If
End Sub


'******************************
'�w�肵���l���������A��v����Z����ԋp
'����  �F
'  lookAtType - �������@�ixlWhole�܂���xlPart�j
'�߂�l�F�Ȃ�
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
    searchValues = CellsValueToArray() ' �I�����ꂽ�Z���̒l��z��Ɋi�[
    If IsEmpty(searchValues) Then
        Exit Function ' ���̏ꍇ�A�����ŏ������I��
    End If

    ' �A�N�e�B�u�V�[�g�̑S�Z��������
    For i = LBound(searchValues) To UBound(searchValues)
        Set foundCell = ws.Cells.Find(What:=searchValues(i), After:=ws.Cells(1, 1), LookIn:=xlValues, _
                                      LookAt:=lookAtType, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                      MatchCase:=False)
        If Not foundCell Is Nothing Then
            firstAddress = foundCell.Address
            Do
                ' �����Ō��������Z����͈͂ɒǉ�
                If combinedRange Is Nothing Then
                    Set combinedRange = foundCell
                Else
                    Set combinedRange = Union(combinedRange, foundCell)
                End If
                Set foundCell = ws.Cells.FindNext(foundCell)
            Loop While Not foundCell Is Nothing And foundCell.Address <> firstAddress
        End If
    Next i

    ' ���������Z��������ΑI�����A�֐��̖߂�l�Ƃ��Đݒ�
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
        ' �֐��̌��ʂ��`�F�b�N
    If IsEmpty(myArray) Then
        ' �֐����牽���Ԃ���Ȃ������ꍇ�̏���
        MsgBox "�����Ԃ���܂���ł����B", vbExclamation
        Exit Sub ' ���̏ꍇ�A�����ŏ������I��
    End If
    ' myArray���g��������...
    ' �z��̓��e��ʒm
    toastMessage = ""
    For i = LBound(myArray) To UBound(myArray)
        toastMessage = toastMessage & myArray(i) & " "
    Next i
    
    ' �g�[�X�g���b�Z�[�W��\��
    ShowToast toastMessage
End Sub

Function CellsValueToArray() As Variant
    Dim selectedRange As Range
    Dim cell As Range
    Dim resultList As New Collection
    Dim resultArray() As Variant
    Dim i As Long

    ' �I������Ă���Z���͈͂��擾
    Set selectedRange = Selection
    
    ' �e�Z���̒l���R���N�V�����Ɋi�[
    For Each cell In selectedRange
        If Not IsEmpty(cell.value) Then
            resultList.Add cell.value
        End If
    Next cell
    
    ' resultList����łȂ����Ƃ��m�F
    If resultList.Count = 0 Then
        MsgBox "��̃Z���݂̂��I������Ă��܂��B", vbExclamation
        Exit Function ' �����Ŋ֐��𔲂���
    End If
    
    ' �R���N�V��������z��ɕϊ�
    ReDim resultArray(1 To resultList.Count)
    For i = 1 To resultList.Count
        resultArray(i) = resultList(i)
    Next i
    
    ' �z���߂�l�Ƃ��ĕԋp
    CellsValueToArray = resultArray
End Function

'******************************
'�I�𒆂̃Z���̍s�����񕝎�������
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �s�����񕝎�������_EUM()
    Selection.Rows.AutoFit
    Selection.Columns.AutoFit
End Sub

'******************************
'�I�𒆂̃Z���̒l����ɐ}�`���쐬
'����  �F�Ȃ�
'�߂�l�F�Ȃ�
'******************************
Sub �}�`�쐬bySelectedCellValue_EUM()
    Dim cell As Range
    Dim shp As Shape
    Dim currentPrefix As String
    Dim leftPosition As Double
    Dim topPosition As Double
    Dim DataRange As Range
    
    ' �I�𒆂̃A�N�e�B�u�Z���͈͂��w��
    Set DataRange = Selection
    
    ' �����l�Ƃ��ċ�̃v���t�B�b�N�X��ݒ�
    currentPrefix = ""
    
    ' �f�[�^�̃Z�������[�v
    For Each cell In DataRange
        If cell.Column <> 1 Then
            If cell.Offset(0, -1).value <> "" Then
                ' �v���t�B�b�N�X���X�V
                currentPrefix = cell.Offset(0, -1).value & "-"
            End If
        End If
        
        If cell.value <> "" Then
            ' �}�`�̍�����̍��W��ϐ���
            leftPosition = cell.Left
            topPosition = cell.Top
            ' �}�`���쐬���e�L�X�g��ݒ�
            ' �}�`�̍쐬
            Set shp = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, leftPosition, topPosition, 137.25, 33)
            ' �h��Ԃ��̐ݒ�
            With shp.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(255, 0, 0)
                .Transparency = 0.8
                .Solid
            End With
            
            ' �g���̐ݒ�
            With shp.Line
                .Visible = msoTrue
                .DashStyle = msoLineSysDash
                .Weight = 1.5
                .ForeColor.RGB = RGB(255, 0, 0)
                .Transparency = 0.5
            End With
            
            ' �e�L�X�g�̐ݒ�
            With shp.TextFrame2.TextRange.Font
                .Size = 16
                .NameComplexScript = "HG�ۺ޼��M-PRO"
                .NameFarEast = "HG�ۺ޼��M-PRO"
                .Name = "HG�ۺ޼��M-PRO"
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
