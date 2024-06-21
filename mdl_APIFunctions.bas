Attribute VB_Name = "mdl_APIFunctions"
Option Compare Database


Private Declare PtrSafe Function apiGetComputerName Lib "kernel32" Alias _
"GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare PtrSafe Function apiGetUserName Lib "advapi32.dll" Alias _
"GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare PtrSafe Function Sleep Lib "kernel32" ( _
ByVal dwMilliseconds As Long)  ' for zipping function


Public Function fOSMachineName() As String
'Returns the computername
Dim lngLen As Long, lngX As Long
Dim strCompName As String
lngLen = 16
strCompName = String$(lngLen, 0)
lngX = apiGetComputerName(strCompName, lngLen)
If lngX <> 0 Then
fOSMachineName = Left$(strCompName, lngLen)
Else
fOSMachineName = ""
End If
End Function
