Option Explicit

Dim Sour, Dest, Ext, Fso, Win, objFSO, objFile, line, filePath
Dim Msg
Sour = "C:\Users\Keita\Desktop\Dev\CopyFilesRecur\Test"        :'�R�s�[���t�H���_
Dest = "C:\Users\Keita\Desktop\Dev\CopyFilesRecur\Target"       :'�R�s�[��t�H���_
filePath = "C:\Users\Keita\Desktop\Dev\CopyFilesRecur\list.txt"         :'�R�s�[�Ώۂ̃t�@�C�������X�g���L�q�����t�@�C��
Set Fso = CreateObject("Scripting.FileSystemObject")
Set Win = Wscript.CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")


'--�ċA�Ăяo���ŃR�s�[�����s
Sub copyFiles(parent, sour, dest)
    Dim folder, fi, fname
    Set folder = Win.NameSpace(sour)
    For Each fi In folder.Items
        If fi.IsFolder Then
            copyFiles fi.Name, sour & "\" & fi.Name, dest
        Else
            Set objFile = objFSO.OpenTextFile(filePath)
            Do Until objFile.AtEndOfStream
                line = objFile.ReadLine
                If line = fi.Name Then
                    '--�R�s�[��Ƀt�@�C�������݂���
                    If Fso.FileExists(Dest & "\" & fi.Name) Then
                        fname = Dest & "\" & parent & fi.Name
                    '--���݂��Ȃ�
                    Else
                        fname = Dest & "\" & fi.Name
                    End If
                    Fso.CopyFile sour & "\" & fi.Name, fname  :'�R�s�[���s
                    Exit Do
                End If
            Loop            '
        End If
    Next
End Sub

'--�R�s�[��t�H���_���Ȃ���΍쐬
If Fso.FolderExists(Dest) Then
Else
    Fso.CreateFolder(Dest)
End If
'--�R�s�[���s
copyFiles "", Sour, Dest


WScript.Echo "�������I�����܂����B"