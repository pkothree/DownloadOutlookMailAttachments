###
# Danny Davis
# twitter: twitter.com/pko3
# github: github.com/pkothree
# Created: 08/22/19
# Modified: 09/09/19
# Description: Get Mail Attachments from Outlook folder
###

# Add assembly and create object to access Outlook
Add-type -assembly "Microsoft.Office.Interop.Outlook"
$olDefaultFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]
$outlook = New-Object -comobject Outlook.Application
 
# Create namespace and get folders
$namespace = $outlook.GetNameSpace("MAPI")
$folders = $namespace.getDefaultFolder($olDefaultFolders::olFolderInBox)

# tests to see im connections work
$namespace.CurrentUser # get the current user
$folders.items # shows all mails
$folders.folder # shows folders
 
# define filepath to save files 
$localpath = "F:\temp\"
# index, not important
$i = 0
# Outlook will pop up and ask you for a folder
# select the folder where your files are
$f = $namespace.PickFolder()
# iterate through every mail in the folder and save the attachements
$f.items | foreach{
    $_.attachments|foreach{
        $i++
        Write-host $i
        Write-Host $_
        $_.saveasfile((Join-Path $localpath $_.filename))
    }
}