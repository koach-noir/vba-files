Attribute VB_Name = "ModuleDevSub"

' �O���[�o���ϐ��Ƃ���CuidGenerator�C���X�^���X��ێ�
Global cuidGenerator As Object

Function EnsureGeneratorInstance() As Object
    If cuidGenerator Is Nothing Then
        ' �C���X�^���X�����������̏ꍇ�̂݁A�V�����C���X�^���X���쐬
        Set cuidGenerator = CreateObject("DevSupLibrary.CuidGenerator")
    End If
    Set EnsureGeneratorInstance = cuidGenerator
End Function

'Function GenerateCUID() As String
'    ' CuidGenerator�C���X�^���X���m���Ɏ擾
'    Dim gen As Object
'    Set gen = EnsureGeneratorInstance()
'    ' �������ꂽCUID��Ԃ�
'    GenerateCUID = gen.GenerateCUID()
'End Function

Sub TestGenerateCUID()
    ' CuidGenerator�N���X�̃C���X�^���X���쐬
    Dim generator As Object
    Set generator = CreateObject("DevSupLibrary.CuidGenerator")

    ' CUID�𐶐�
    Dim cuid As String
    cuid = generator.GenerateCUID()
    Debug.Print cuid  ' �C�~�f�B�G�C�g�E�B���h�E��CUID���o��
End Sub


Sub DeleteSheetsExceptSheet1()
    Dim ws As Worksheet
    Dim wsToDelete As Worksheet

    ' �V�[�g1�ȊO�̂��ׂẴV�[�g���폜
    Application.DisplayAlerts = False ' �폜�̊m�F�_�C�A���O��\�����Ȃ��悤�ɐݒ�
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Sheet1" Then
            'ws.Delete ' �����̖h�~CO
        End If
    Next ws
    Application.DisplayAlerts = True ' �f�t�H���g�̐ݒ�ɖ߂�
End Sub
