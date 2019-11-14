Option Explicit

Const vbNormalFocus = 1
Const KBNoArray = "KB4484127,KB4484119,KB4484113" 'KB�ԍ��̔z��
Const SubKeyNameForSameBit = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
Const SubKeyNameForDifferentBit = "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

Dim cmd
Dim KBNoList
KBNoList = Split(KBNoArray, ",")
Dim KBNo

'32 �r�b�g Windows ��� 32 �r�b�g Office/64 �r�b�g Windows ��� 64 �r�b�g Office�̏ꍇ
For Each KBNo In KBNoList
  cmd = "" '������
  cmd = GetUninstallString(KBNo, SubKeyNameForSameBit)
  If Len(Trim(cmd)) > 0 Then
    CreateObject("WScript.Shell").Run cmd, vbNormalFocus, False
    Sleep 10000
  Else
    WScript.Echo "�w�肵��KB�ԍ�[" & KBNo & "]��[UninstallString]���擾�ł��܂���ł����B"
  End If
Next

'64 �r�b�g Windows ��� 32 �r�b�g Office�̏ꍇ
For Each KBNo In KBNoList
  cmd = "" '������
  cmd = GetUninstallString(KBNo, SubKeyNameForDifferentBit)
  If Len(Trim(cmd)) > 0 Then
    CreateObject("WScript.Shell").Run cmd, vbNormalFocus, False
    Sleep 10000
  Else
    WScript.Echo "�w�肵��KB�ԍ�[" & KBNo & "]��[UninstallString]���擾�ł��܂���ł����B"
  End If
Next

Public Function GetUninstallString(ByVal KBNo, ByVal SubKeyName)
'�w�肵��KB�ԍ���[UninstallString]�����W�X�g������擾
  Dim ret
  Dim reg
  Dim names
  Dim display_name
  Dim uninstall_string
  Dim i

  Const HKEY_LOCAL_MACHINE = &H80000002

  ret = "" '������
  Set reg = CreateObject("WbemScripting.SWbemLocator") _
            .ConnectServer(, "root\default") _
            .Get("StdRegProv")
  reg.EnumKey HKEY_LOCAL_MACHINE, SubKeyName, names
  If Not IsNull(names) Then
    On Error Resume Next
    For i = LBound(names) To UBound(names)
      display_name = ""
      reg.GetStringValue HKEY_LOCAL_MACHINE, _
                         SubKeyName & ChrW(92) & names(i), _
                         "DisplayName", _
                         display_name
      '[DisplayName]��KB�ԍ����܂܂�Ă��邩����
      If InStr(LCase(display_name), LCase(KBNo)) Then
        uninstall_string = ""
        reg.GetStringValue HKEY_LOCAL_MACHINE, _
                           SubKeyName & ChrW(92) & names(i) & ChrW(92), _
                           "UninstallString", _
                           uninstall_string
        ret = uninstall_string
        Exit For
      End If
    Next
    On Error GoTo 0
  End If
  GetUninstallString = ret
End Function
