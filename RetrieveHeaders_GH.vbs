'* Refinement of v2
'* Split the header into an array for reading
'* rather than re-reading line by line....slow
'* back to writing to a text file


On Error Resume Next

Const PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
Const FOR_WRITING = 2
GetPublicFolder ("Folder1\SubFolder")


Function GetPublicFolder(strFolderPath)

Dim arrFolders
Dim objOLApp, objOLNS, objFolder
Dim colFolder
Dim i

Dim objFS, objFile
Dim strFileName

Dim c, k
Dim objItem, objPA
Dim strIHead, strArray
Dim myStr, evalStr, writeStr


strFolderPath = Replace(strFolderPath, "/", "\")
arrFolders = Split(strFolderPath, "\")

Set objOLApp = CreateObject("Outlook.Application")
Set objOLNS = objOLApp.GetNamespace("MAPI")
'objOLSession.Logon "", "", False, False

Set objFolder = objOLNS.GetDefaultFolder(18)
Set objFolder = objFolder.Folders.Item(arrFolders(0))
If Not objFolder Is Nothing Then
    For i = 1 To UBound(arrFolders)
        Set colFolders = objFolder.Folders
        Set objFolder = Nothing
        Set objFolder = colFolders.Item(arrFolders(i))
        If objFolder Is Nothing Then
            Exit For
        End If
    Next
End If

strFileName = "C:\Temp\header3.txt"
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFile = objFS.OpenTextFile(strFileName,FOR_WRITING)
writeStr = ""
c = objfolder.Items.Count
For i = 1 To c
	Set objItem = objFolder.Items(i)
	Set objPA = objitem.PropertyAccessor
	strIHead = objPA.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS)	
	strArray = Split(strIHead,vbNewLine)
	For k = LBound(strArray) To UBound(strArray)
		myStr = Trim(strArray(k))
		evalStr = Left(myStr,4)
		If evalStr = "auth" Or evalStr = "smtp" Or evalStr = "Subj" Or evalStr = "Date" Then
			objFile.WriteLine myStr
		End If
	Next
		
	objFile.WriteLine String(100,"=")

Next

objFile.Close

'Set GetPublicFolder = objfolder
'WScript.Echo objfolder.Items.Count
Set colFolders = Nothing
Set objApp = Nothing
Set objFile = Nothing
Set objFS = Nothing
Set objFolder = Nothing
Set objItem = Nothing
Set objPA = Nothing


End Function
