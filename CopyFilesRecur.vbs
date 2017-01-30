Option Explicit

Dim Sour, Dest, Ext, Fso, Win, objFSO, objFile, line, filePath
Dim Msg
Sour = "C:\Users\Keita\Desktop\Dev\CopyFilesRecur\Test"        :'コピー元フォルダ
Dest = "C:\Users\Keita\Desktop\Dev\CopyFilesRecur\Target"       :'コピー先フォルダ
filePath = "C:\Users\Keita\Desktop\Dev\CopyFilesRecur\list.txt"         :'コピー対象のファイル名リストを記述したファイル
Set Fso = CreateObject("Scripting.FileSystemObject")
Set Win = Wscript.CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")


'--再帰呼び出しでコピーを実行
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
                    '--コピー先にファイルが存在する
                    If Fso.FileExists(Dest & "\" & fi.Name) Then
                        fname = Dest & "\" & parent & fi.Name
                    '--存在しない
                    Else
                        fname = Dest & "\" & fi.Name
                    End If
                    Fso.CopyFile sour & "\" & fi.Name, fname  :'コピー実行
                    Exit Do
                End If
            Loop            '
        End If
    Next
End Sub

'--コピー先フォルダがなければ作成
If Fso.FolderExists(Dest) Then
Else
    Fso.CreateFolder(Dest)
End If
'--コピー実行
copyFiles "", Sour, Dest


WScript.Echo "処理が終了しました。"