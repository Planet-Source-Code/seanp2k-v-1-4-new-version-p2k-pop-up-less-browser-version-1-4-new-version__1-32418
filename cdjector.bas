Attribute VB_Name = "CDjector"
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Function openCD(ByVal dRv As String) As Long
Dim Alias As String
Dim retval As Long
Alias = "Drive" & dRv
retval = -1           'we need to set retval to anything other then 0
retval = mciSendString("open " & dRv & ": type cdaudio alias " & Alias & " wait", vbNullString, 0&, 0&)
retval = mciSendString("set " & Alias & " door open", vbNullString, 0&, 0&)
openCD = retval
End Function
Public Function closeCD(ByVal dRv As String) As Long
Dim Alias As String
Dim retval As Long
Alias = "Drive" & dRv
retval = -1           'we need to set retval to anything other then 0
 retval = mciSendString("set " & Alias & " door closed", vbNullString, 0&, 0&)
 retval = mciSendString("close " & Alias, vbNullString, 0&, 0&)
closeCD = retval
End Function



