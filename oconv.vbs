'--------------------------------------
'Convert xlsx and docx into xls,doc
'James Yang <jamesyang999@gmail.com>
'Based on script:
'Script to convert .doc to .docx files
'16.6.2011 FNL
'--------------------------------------
bRecursive = False
sFolder = "D:\Test"
Set oFSO = CreateObject("Scripting.FileSystemObject")

Set oWord = CreateObject("Word.Application")
oWord.Visible = False
oWord.DisplayAlerts = False

Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = False
oExcel.DisplayAlerts = False

Set oFolder = oFSO.GetFolder(sFolder)
ConvertFolder(oFolder)

oWord.Quit
oExcel.Quit

Sub deleteFile(oFile)
    On Error Resume Next
    ' Delete file using del command
    WScript.CreateObject("WScript.Shell").Run "cmd.exe /C del /f /q """ & sFolder & "\" & oFile.name & """", 1, True
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

Sub ConvertFolder(oFldr)
    On Error Resume Next

    For Each oFile In oFldr.Files

        If Left(oFile.Name, 2) = "~$" Then
           ' Skip backup files
        Else

        If LCase(oFSO.GetExtensionName(oFile.Name)) = "docx" Then
            Set oDoc = oWord.Documents.Open(oFile.path)
            oWord.ActiveDocument.SaveAs regReplace("x$", oFile.path, ""), 0    '0=Word97 Document
            oDoc.Close
            deleteFile oFile
        End If

        If LCase(oFSO.GetExtensionName(oFile.Name)) = "xlsx" Then
            Set wb = oExcel.Workbooks.Open(oFile.path)
            oExcel.ActiveWorkbook.SaveAs regReplace("x$", oFile.path, ""), 56   '56=Excel8=Excel 97-2003
            oExcel.ActiveWorkbook.Close
            deleteFile oFile
        End If

        End If
    Next

    If bRecursive Then
        For Each oSubfolder In oFldr.Subfolders
            ConvertFolder oSubfolder
        Next
    End If
End Sub

