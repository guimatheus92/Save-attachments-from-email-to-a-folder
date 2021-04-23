'------------------------------------ Global variables --------------------------------------'
Public DiretorioAnexos As String
Public WithEvents olItems As Outlook.Items
Public responder As Integer
Public MyDir As String, fn As String

' ------------------------------------------------------------------------------------------- '
' Title: Debug Macros
' Author: Guilherme Matheus
' Date: Script created on 23.04.2021
' Script and data info: This script can debug macros that are being called from rules, that can`t be debuged here
'--------------------------------------------------------------------------------------------'

Sub Execute_Macro_ForEach_Email()
    Dim x, mailItem As Outlook.mailItem
    For Each x In Application.ActiveExplorer.Selection
    DoEvents
        If TypeName(x) = "MailItem" Then
            Set mailItem = x
            DoEvents
            'Call a macro to debug
            Call Save_Attachments(mailItem)
            DoEvents
        End If
    DoEvents
    Next
End Sub


' ------------------------------------------------------------------------------------------- '
' Title: Save e-mail attachments
' Author: Guilherme Matheus
' Date: Script created on 23.04.2021
' Script and data info: This script can save e-mail attachments according to file extension
'--------------------------------------------------------------------------------------------'

Sub Save_Attachments(Email As Outlook.mailItem)
                    
    'Dim strSubject As String
    Dim objMsg As Outlook.mailItem
    Dim objSubject As String
    
    objSubject = Email.Subject
    
    Dim oItem As Object
    Dim propertyAccessor As Outlook.propertyAccessor
    Set oItem = Application.ActiveExplorer.Selection.Item(1)
    Set propertyAccessor = oItem.propertyAccessor
    status_email = propertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10810003")
    Set oItem = Nothing
    
    'I check if the email has already been answered or forwarded
    'Email status:
    If status_email <> 0 Then Exit Sub
    
        'I define which path will be saved the files in attachments of the emails
        DiretorioAnexos = "\\VMSPFSFSCH15\LDH_Share\automaticarchive\staging\prod\decrypted"
    
        'Debug.Print DiretorioAnexos
    
        'Debug.Print objSubject
    
        Dim MailID As String
        Dim Mail As Outlook.mailItem

        MailID = Email.EntryID
        Set Mail = Application.Session.GetItemFromID(MailID)
    
        'I check if the attachment in the email is a single CSV file and save all
        For Each Anexo In Mail.Attachments
            If Right(Anexo.FileName, 3) = "csv" Then
                Anexo.SaveAsFile DiretorioAnexos & "\" & Anexo.FileName
            End If
        Next
    
        'I check if the attachment in the email is a single ZIP file and save all
        For Each Anexo In Mail.Attachments
            If Right(Anexo.FileName, 3) = "zip" Then
                Anexo.SaveAsFile DiretorioAnexos & "\" & Anexo.FileName
                Call Unzip_Files
            End If
        Next
        DoEvents
    
        Set Mail = Nothing
            
 End Sub

' ------------------------------------------------------------------------------------------- '
' Title: Unzip a zipped file
' Author: Guilherme Matheus
' Date: Script created on 23.04.2021
' Script and data info: This script can unzip a zipped file through Windows object
'--------------------------------------------------------------------------------------------'

Sub UnzipAFile(zippedFileFullName As Variant, unzipToPath As Variant)

    Dim ShellApp As Object

    'Copy the files & folders from the zip into a folder
    Set ShellApp = CreateObject("Shell.Application")
    'Number 4 and 16 is a flag that don't show the msg box in case the file already exists and accept the option "Yes to All"
    ShellApp.NameSpace(unzipToPath).CopyHere ShellApp.NameSpace(zippedFileFullName).Items, 4 + 16

End Sub

' ------------------------------------------------------------------------------------------- '
' Title: Unzip all zipped files on the destination folders
' Author: Guilherme Matheus
' Date: Script created on 31.03.2021
' Script and data info: This script can unzip a zipped file from the destionation folder when called the "UnzipAFile" macro through Windows object
'--------------------------------------------------------------------------------------------'

Sub Unzip_Files()

    Dim path_file As Variant
    Dim path_ext As Variant
    Dim file_name As String
    
    'Path that the file will be unziped
    path_ext = "\\VMSPFSFSCH15\LDH_Share\automaticarchive\staging\prod\decrypted\"
    
    'Find out the name of the zip file + the path it is in
    file_name = Dir(path_ext & "*.zip")
        
    'Path that the zipped file is located
    path_file = "\\VMSPFSFSCH15\LDH_Share\automaticarchive\staging\prod\decrypted\" & file_name
        
    'Run the macro as long as there are zipped files in the folder
    Do While Len(file_name) > 0
        
        'The first part refers to the name of the file that will be unziped, and the second part refers to the path that it will be unziped.
        Call UnzipAFile(path_file, path_ext)
        
        'I delete the first zip file that was extracted
        'First removes the "read-only" file attribute if set
        On Error Resume Next
        SetAttr FileToDelete, vbNormal
        'Then I delete the file
        Kill path_file
        
        'Search for the next file
        file_name = Dir
    
        'Displays success message
        'MsgBox "Arquivo " & file_name & "descompactados! " & "no diretório: " & path_ext
        
    Loop

End Sub
