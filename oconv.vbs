'--------------------------------------
'Convert xlsx and docx into xls,doc
'James Yang <jamesyang999@gmail.com>
'--------------------------------------
Set args = Wscript.Arguments

If args.count < 1 Then Wscript.Quit

Set oFSO = CreateObject("Scripting.FileSystemObject")

Set oWord = CreateObject("Word.Application")
oWord.Visible = False
oWord.DisplayAlerts = False

Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = False
oExcel.DisplayAlerts = False

ConvertFile(args(0))

oWord.Quit
oExcel.Quit

Sub deleteFile(filePath)
    On Error Resume Next
    ' Delete file using del command
    WScript.CreateObject("WScript.Shell").Run "cmd.exe /C del /f /q """ & filePath & """", 1, True
End Sub

Function regReplace(patrn, str1, replStr)
  Dim regEx

  ' Create regular expression.
  Set regEx = New RegExp
  regEx.Pattern = patrn
  regEx.IgnoreCase = True

  ' Make replacement.
  regReplace = regEx.Replace(str1, replStr)
End Function

Sub ConvertFile(filePath)
    On Error Resume Next

    Dim fileName, saveFolder
    fileName = oFSO.GetFileName(filePath)

    If args.count>1 Then
       saveFolder = args(1)
    else
       saveFolder = oFSO.GetParentFolderName(filePath) & "\ok"
    End If

    oFSO.CreateFolder saveFolder

    If Left(fileName, 2) = "~$" Then
       ' Skip backup files
    Else

    If LCase(oFSO.GetExtensionName(fileName)) = "docx" Then
        Set oDoc = oWord.Documents.Open(filePath)
        oWord.ActiveDocument.SaveAs saveFolder & "\" & regReplace("x$", fileName, ""), 0    '0=Word97 Document
        oDoc.Close
        deleteFile filePath
    End If

    If LCase(oFSO.GetExtensionName(fileName)) = "xlsx" Then
        Set wb = oExcel.Workbooks.Open(filePath)
        oExcel.ActiveWorkbook.SaveAs saveFolder & "\" & regReplace("x$", fileName, ""), 56   '56=Excel8=Excel 97-2003
        oExcel.ActiveWorkbook.Close
        deleteFile filePath
    End If

    End If

End Sub

