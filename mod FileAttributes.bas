Attribute VB_Name = "modFileAttributes"
'Call SetFileAttributes("C:\Windows\Desktop\Test.txt", True, True, False)
'MsgBox GetFileAttribute("C:\Windows\Desktop\Test.txt", Archive)
'MsgBox GetFileAttribute("C:\Windows\Desktop\Test.txt", Hidden)
'MsgBox GetFileAttribute("C:\Windows\Desktop\Test.txt", ReadOnly)

Public Enum FileAttributes
    Archive
    Hidden
    ReadOnly
End Enum

Public Function SetFileAttributes(TheFile As String, ReadOnly As Boolean, Hidden As Boolean, Archive As Boolean)
    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set Att = Fso.GetFile(TheFile)
    If ReadOnly = True And Hidden = False And Archive = False Then Att.Attributes = 1
    If ReadOnly = True And Hidden = True And Archive = False Then Att.Attributes = 3
    If ReadOnly = True And Hidden = True And Archive = True Then Att.Attributes = 35
    If ReadOnly = True And Hidden = False And Archive = True Then Att.Attributes = 33
    If ReadOnly = False And Hidden = True And Archive = True Then Att.Attributes = 34
    If ReadOnly = False And Hidden = True And Archive = False Then Att.Attributes = 2
    If ReadOnly = False And Hidden = False And Archive = True Then Att.Attributes = 32
    If ReadOnly = False And Hidden = False And Archive = False Then Att.Attributes = 0
End Function

Public Function GetFileAttribute(TheFile As String, TheAttribute As FileAttributes) As Boolean
    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set Att = Fso.GetFile(TheFile)
    
    If TheAttribute = Archive Then
        If (Att.Attributes = 32) Or (Att.Attributes = 33) Or (Att.Attributes = 34) Or (Att.Attributes = 35) Then
            GetFileAttribute = True
        Else
            GetFileAttribute = False
        End If
        Exit Function
    End If
    
    If TheAttribute = Hidden Then
        If (Att.Attributes = 2) Or (Att.Attributes = 3) Or (Att.Attributes = 34) Or (Att.Attributes = 35) Then
            GetFileAttribute = True
        Else
            GetFileAttribute = False
        End If
        Exit Function
    End If
    
    If TheAttribute = ReadOnly Then
        If (Att.Attributes = 1) Or (Att.Attributes = 3) Or (Att.Attributes = 33) Or (Att.Attributes = 35) Then
            GetFileAttribute = True
        Else
            GetFileAttribute = False
        End If
        Exit Function
    End If
End Function
