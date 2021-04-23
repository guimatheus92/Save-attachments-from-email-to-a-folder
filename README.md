# Save attachments from e-mail to a folder

This repository will be used to store the VBA scripts for Microsoft Outlook run a macro that can save any attachments from email to a folder.

## **Prerequisites**

***Microsoft Outlook***

For VBA scripts in Microsoft Outlook to work, you must check the following points below:
    
*  Using a mailbox for any user that has access to the email

*  Check if the rules on Outlook can run a script. If not, run the regedit file to add the option `Run a Script`

*  Necessary to enable All Macros from Trust Settings on MS Outlook

*  Create a rule to execute the VBA macro on the e-mail received

***Installation***

If your Microsoft Outlook does not have the `Run a Script` option in the rule, then you must add a value within a parameter in Regedit or just execute the regedit file.

But it is easier if you download the file with the extension `.reg` and run it, according to the version of your Microsoft Outlook. After installation, restart Microsoft Outlook.
That way, the `Run a Script` option will appear normally.

1.  Download the regedit file according to the version of your Microsoft Outlook
2.  Execute the regedit file to add value in the right path
3.  Restart your Microsoft Outlook

***In order to work, you need to copy and paste the content of the file `ThisOutlookSession.vb` into MS Outlook.***

***Outlook Rule***

In order to create your rule on Outlook, you can import the rule called `SaveAtchFromEmail.rwz` to your MS Outlook rules and change the email subject that the file will be received as the image in the repository.

You can access the article on [Medium](https://guimatheus92.medium.com/save-attachments-from-e-mail-to-a-folder-with-ms-outlook-vba-2382d3917abc "Medium") as well!
