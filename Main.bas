Attribute VB_Name = "Main"
Option Explicit

Private Declare PtrSafe Function sqlite3_open Lib "winsqlite3.dll" (ByVal filename As LongPtr, ByRef ppDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_exec Lib "winsqlite3.dll" (ByVal db As LongPtr, ByVal sql As LongPtr, ByVal callback As LongPtr, _
                                                                    ByVal param As LongPtr, ByRef errmsg As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_close Lib "winsqlite3.dll" (ByVal db As LongPtr) As Long

Const DBPATH = "z:\test.db"

Sub ExecuteQuery()
    Dim db As LongPtr
    Dim Retval As Long
    Dim strSQL As String
    Dim errmsg As LongPtr
    
    ' �f�[�^�x�[�X���J��
    Retval = sqlite3_open(StrPtr(StringToUTF8(DBPATH)), db)
    If Retval <> 0 Then
        Call MsgBox("�f�[�^�x�[�X�̃I�[�v���ɂɎ��s���܂���")
        Exit Sub
    End If
    
    ' �e�[�u�����쐬����SQL��
    strSQL = "CREATE TABLE IF NOT EXISTS �e�X�g ( " & _
          "ID INTEGER PRIMARY KEY, " & _
          "���O TEXT NOT NULL, " & _
          "���w INTEGER, " & _
          "�p�� INTEGER, " & _
          "���� INTEGER);"

    Retval = sqlite3_exec(db, StrPtr(StringToUTF8(strSQL)), 0, 0, errmsg)
    If Retval <> 0 Then
        Call MsgBox("�e�[�u���̍쐬�Ɏ��s���܂���: " & UTF8ToString(errmsg))
        Call sqlite3_close(db)
        Exit Sub
    End If

    ' �f�[�^��}������SQL��
    strSQL = "INSERT INTO �e�X�g (���O, ���w, �p��, ����) VALUES " & _
          "('���', 85, 90, 88), " & _
          "('����', 78, 82, 85), " & _
          "('�c��', 92, 88, 100);"

    ' SQL�����s���ăf�[�^��}��
    Retval = sqlite3_exec(db, StrPtr(StringToUTF8(strSQL)), 0, 0, errmsg)
    If Retval <> 0 Then
        Call MsgBox("�f�[�^�̑}���Ɏ��s���܂���: " & UTF8ToString(errmsg))
        Call sqlite3_close(db)
        Exit Sub
    End If
    
    Call sqlite3_close(db)
End Sub

