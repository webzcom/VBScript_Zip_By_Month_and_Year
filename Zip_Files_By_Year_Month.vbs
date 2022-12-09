'File System Object Prep
Const ForReading = 1
Const ForWriting = 2

sFolder = InputBox("Enter log folder path:","Select a Log Folder to Compress","C:\inetpub\logs\LogFiles\W3SVC3")
Set oFSO = CreateObject("Scripting.FileSystemObject")

For Each oFile In oFSO.GetFolder(sFolder).Files

	on error resume next
	'Breakdown file name
    strFileType = Right(oFile.Name,3)

	if strFileType = "csv" then
		strTemp = Replace(Mid(oFile.Name,20,Len(oFile.Name)-4),".csv","")
		arrDate = Split(strTemp,"_")
		iYear = Left(arrDate(0),2)
		iMonth = arrDate(1)
		if Len(iMonth) < 2 then
			iMonth = "0" & iMonth
		end if
		CheckValue = arrDate(1)
		CurrentMonth = Mid(DatePart("yyyy", Now()),3,2) & DatePart("m", Now())

        if iYear & iMonth = CurrentMonth and (strFileType = "log" OR strFileType = "csv")	then
            'Do not process current month file, only archive previous months
             'msgbox("Skipping " & sFolder & "\" & oFile.Name)
        else
            WindowsZip sFolder & "\" & oFile.Name, sFolder & "\" & iYear & iMonth & ".zip"
		end if
	end if
 
    if strFileType = "log" then	
		iYear = Mid(oFile.Name,5,2)
		iMonth = Mid(oFile.Name, 7,2)
		CheckValue = iYear & iMonth
		CurrentMonth = Mid(DatePart("yyyy", Now()),3,2) & DatePart("m", Now())

        if iYear & iMonth = CurrentMonth and (strFileType = "log" OR strFileType = "csv")	then
            'Do not process current month file, only archive previous months
            'msgbox("Skipping " & sFolder & "\" & oFile.Name)
        else
            WindowsZip sFolder & "\" & oFile.Name, sFolder & "\" & iYear & iMonth & ".zip"
        end if
	end if

Next

Function WindowsUnZip(sUnzipFileName, sUnzipDestination)
 'This script is provided under the Creative Commons license located
  'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
  'be used for commercial purposes with out the expressed written consent
  'of NateRice.com

  Set oUnzipFSO = CreateObject("Scripting.FileSystemObject")
  If Not oUnzipFSO.FolderExists(sUnzipDestination) Then
    oUnzipFSO.CreateFolder(sUnzipDestination)
  End If

  With CreateObject("Shell.Application")
       .NameSpace(sUnzipDestination).Copyhere .NameSpace(sUnzipFileName).Items
  End With

  Set oUnzipFSO = Nothing
End Function

'To Test Windows Zip Function Separately 
'WindowsZip "C:\test\test2.txt","C:\test\test.zip"

Function WindowsZip(sFile, sZipFile)
  'This script is provided under the Creative Commons license located
  'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
  'be used for commercial purposes with out the expressed written consent
  'of NateRice.com

  Set oZipShell = CreateObject("WScript.Shell") 
  Set oZipFSO = CreateObject("Scripting.FileSystemObject")

  If Not oZipFSO.FileExists(sZipFile) Then
    NewZip(sZipFile)
  End If


  Set oZipApp = CreateObject("Shell.Application")
  sZipFileCount = oZipApp.NameSpace(sZipFile).items.Count
  aFileName = Split(sFile, "\")
  sFileName = (aFileName(Ubound(aFileName)))

  'listfiles
  sDupe = False

  For Each sFileNameInZip In oZipApp.NameSpace(sZipFile).items
    If LCase(sFileName) = LCase(sFileNameInZip) Then
      sDupe = True
      Exit For
    End If
  Next
 
  If Not sDupe Then
    oZipApp.NameSpace(sZipFile).Copyhere sFile
    'Keep script waiting until Compressing is done
    On Error Resume Next
    sLoop = 0
    Do Until sZipFileCount < oZipApp.NameSpace(sZipFile).Items.Count
      Wscript.Sleep(100)
      sLoop = sLoop + 1
    Loop
    On Error GoTo 0
  End If
End Function

Sub NewZip(sNewZip)
  'This script is provided under the Creative Commons license located
  'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
  'be used for commercial purposes with out the expressed written consent
  'of NateRice.com

  Set oNewZipFSO = CreateObject("Scripting.FileSystemObject")
  Set oNewZipFile = oNewZipFSO.CreateTextFile(sNewZip)

  oNewZipFile.Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0)
  oNewZipFile.Close

  Set oNewZipFSO = Nothing
  Wscript.Sleep(500)
End Sub
