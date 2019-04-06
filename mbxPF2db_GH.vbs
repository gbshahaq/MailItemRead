'Const PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
TransferPublicFolder ("Folder1\SubFolder")



Function TransferPublicFolder(strFolderPath)

strFolderPath = Replace(strFolderPath, "/", "\")
arrFolders = Split(strFolderPath, "\")

'Setup mailbox connection and locate folder
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

'setup database connection and establish recordset
Const adOpenDynamic = 2
Const adLockOptimistic = 3
Const olFolderInbox = 6
Const adOpenKeyset = 1
Const adLockBatchOptimistic = 4

'hard coded db server, database - use Windows Authentication
strConn = "Provider = sqloledb; Data Source = dbserver; Initial Catalog = myDB; Integrated Security = SSPI;"
'use ADO for database object - good performance and simple syntax
Set cnS = CreateObject("ADODB.Connection")
Set rsS = CreateObject("ADODB.Recordset")

'open connection to SQL DB
cnS.Open strConn

rsS.Open "SELECT * FROM dbo.tbl_PFMailLog " , cnS, adOpenDynamic, adLockOptimistic

c = objfolder.Items.Count

For i = 1 To c
	Set objItem = objFolder.Items(i)
	With objItem
	
		MailID = 		.EntryID
		DateSent = 		.ReceivedTime
		FromAddress = 	.senderemailaddress
		Subj = 			.Subject
'		Set objPA = .PropertyAccessor
'		strIHead = objPA.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS)
	End With
		
    rsS.AddN
        rsS.Fields("MailID") = MailID
        rsS.Fields("DateSent") = DateSent
        rsS.Fields("Sender") = FromAddress
        rsS.Fields("Subject") = Subj
    rsS.Update

Next


Set colFolders = Nothing
Set objOLApp = Nothing
Set objFolder = Nothing
Set objItem  = Nothing
Set objOLNS = nothing

rss.Close
cns.Close
Set rss = Nothing
Set cns = Nothing



End Function
