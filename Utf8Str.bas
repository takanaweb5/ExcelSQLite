Attribute VB_Name = "Utf8Str"
Option Explicit

Private Const CP_UTF8 As Long = 65001
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, _
                                                                     ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long

Public Function UTF8ToString(ByVal pUtf8 As LongPtr) As String
On Error GoTo ErrHandle
    Dim size As Long
    size = MultiByteToWideChar(CP_UTF8, 0, pUtf8, -1, 0, 0)
    'size��null�I�[�������܂ނ��߁A����0�̕�����̎���1���Ԃ�
    If size <= 1 Then Exit Function
    
    Dim result As String
    '�K�v�Ȓ������m�ۂ���(null�I�[������)
    result = String(size - 1, vbNullChar)
    If MultiByteToWideChar(CP_UTF8, 0, pUtf8, -1, StrPtr(result), size) <> 0 Then
        UTF8ToString = result
    End If
    
    Exit Function
ErrHandle:
    UTF8ToString = ""
End Function

Public Function StringToUTF8(ByVal text As String) As String
On Error GoTo ErrHandle
    If text = "" Then Exit Function
    
    Dim size As Long
    size = WideCharToMultiByte(CP_UTF8, 0, StrPtr(text), -1, 0, 0, 0, 0)
    If size = 0 Then Exit Function
    
    ReDim buf(size) As Byte
    If WideCharToMultiByte(CP_UTF8, 0, StrPtr(text), -1, VarPtr(buf(0)), size, 0, 0) <> 0 Then
        StringToUTF8 = buf
    End If
    
    Exit Function
ErrHandle:
    StringToUTF8 = ""
End Function

