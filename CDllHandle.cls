VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDllHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PATHの通っていないsqlite3.dllをDeclare宣言で使用可能にするためのクラス

Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long

Private Const DLLPATH = "C:\Fujitsu\Softwaredisc\Software2\ODD\Software\Application\PowerDirector_AVCHD\bin\powerdirector\Setup\Utility\InteropPalette\x64\Lib\sqlite3.dll"
Private hDllHandle As LongPtr

'コンストラクタ
Private Sub Class_Initialize()
    hDllHandle = LoadLibrary(DLLPATH)
End Sub

'デストラクタ
Private Sub Class_Terminate()
    Call FreeLibrary(hDllHandle)
End Sub
