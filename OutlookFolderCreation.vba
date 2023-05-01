Option Explicit

Dim objOutlook, objNS, objFolder
Dim strFolderName, strStartDate
Dim iWeeks, i

'Create an instance of the Outlook application
Set objOutlook = CreateObject("Outlook.Application")
Set objNS = objOutlook.GetNamespace("MAPI")

'Set the starting date for the folders
strStartDate = "01/01/2023"

'Set the number of weeks to create folders for
iWeeks = 48

For i = 1 To iWeeks
    'Create a folder for Monday
    strFolderName = "Monday " & strStartDate
    Set objFolder = objNS.Folders.Add(strFolderName)
    objFolder.Move objNS.Folders("Inbox")

    'Create a folder for Tuesday
    strFolderName = "Tuesday " & strStartDate
    Set objFolder = objNS.Folders.Add(strFolderName)
    objFolder.Move objNS.Folders("Inbox")

    'Create a folder for Wednesday
    strFolderName = "Wednesday " & strStartDate
    Set objFolder = objNS.Folders.Add(strFolderName)
    objFolder.Move objNS.Folders("Inbox")

    'Create a folder for Thursday
    strFolderName = "Thursday " & strStartDate
    Set objFolder = objNS.Folders.Add(strFolderName)
    objFolder.Move objNS.Folders("Inbox")

    'Create a folder for Friday
    strFolderName = "Friday " & strStartDate
    Set objFolder = objNS.Folders.Add(strFolderName)
    objFolder.Move objNS.Folders("Inbox")

    'Increment the start date by 7 days
    strStartDate = DateAdd("d", 7, strStartDate)
Next

'Clean up
Set objFolder = Nothing
Set objNS = Nothing
Set objOutlook = Nothing
