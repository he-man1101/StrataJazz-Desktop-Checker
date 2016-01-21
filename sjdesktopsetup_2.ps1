# Add the required .NET assembly:
Add-Type -AssemblyName System.Windows.Forms

$TimeStamp = get-date -format g

#variables for creating directory to store the log file

$MyDocuments = [environment]::getfolderpath("mydocuments")
$DocFolder = 'Strata Decision Technology'
$DocSubFolder = 'Desktop Requirements Checker'
$DocFolderType = 'directory'
$OutputLoc = "$MyDocuments\$DocFolder\$DocSubFolder\desktopcheckerlog.txt"

#conditional loop delay caused out-file dictorynotfoundexception
new-item -path $MyDocuments -name $DocFolder -itemtype $DocFolderType -ErrorAction SilentlyContinue
new-item -path "$MyDocuments\$DocFolder" -name $DocSubFolder -itemtype $DocFolderType -ErrorAction SilentlyContinue

if((Test-Path -path $MyDocuments) -eq $false)
   {
   #warn user that a log file would not be present since the local documents folder could not be located
   Write-Warning "My Documents file not found. Log file will not be created."
   }
   elseif((test-path -Path "$MyDocuments\$DocFolder") -eq $false)
   {
   ##Create Strata Decision Technology folder under My Documents
   new-item -path $MyDocuments -name $DocFolder -itemtype $DocFolderType
   }
   elseif((test-path -Path "$MyDocuments\$DocFolder\$DocSubFolder") -eq $false)
   {
   ##Create Desktop Requirements Checker Folder under the Strata Decision Technology folder
   new-item -path "$MyDocuments\$DocFolder" -name $DocSubFolder -itemtype $DocFolderType
   }
  


#create a function to check existing registry key
function test-RegistryKey {

param (
 [parameter(Mandatory=$true)]
 [ValidateNotNullOrEmpty()]$Path,
 [parameter(Mandatory=$true)]
 [ValidateNotNullOrEmpty()]$Value
)

 

if((get-item -path $Path).getvalue($Value) -ne $null -eq $true)
 {
    return $true
 }
else{
return $false
}
}

#create a function to set/create a registry value
Function set-RegistryValue{
   Param
   (
   $RegPath,
   $RegName,
   $RegValue,
   $RegType
   )
try{   
    #check registry key
   if((Test-Path -path $RegPath) -eq $false)
   {
   "$TimeStamp $RegPath does not exist." | out-file -filepath $OutputLoc -append
   return
   #exit function since registry path not found
   }
   elseif((test-registrykey -path $RegPath -value $RegName) -eq $false)
   {
   #create a new registry key if it doesn't exist
   new-itemproperty -path $RegPath -name $RegName -value $RegValue -type $RegType -force
   "$TimeStamp Created registry key $RegName." | out-file -filepath $OutputLoc -append
   }
   #check registry value
   elseif((get-itemproperty -path $RegPath).$RegName -ne $RegValue)
   {
   #update registry value if it doesn't match
   set-itemproperty -path $RegPath -name $RegName -value $RegValue -type $RegType -force -erroraction silentlycontinue
   "$TimeStamp Updated registry key $RegName value to $RegValue." | out-file -filepath $OutputLoc -append 
   return
   }
   }
catch{
     $_.Exception.Message -like "* already exists*" #why? why? why?
     }
finally{
      #yay
       "$TimeStamp Registry value $RegValue for registry key named $RegName is already set correctly." | out-file -filepath $OutputLoc -append
    }
   
}
#create folder in registry
Function set-FolderValue{
   Param
   (
   $FolderPath,
   $FolderName,
   $FolderType
   )
try{   
   #check folder path
   if((Test-Path -path $FolderPath) -eq $false)
   {
   "$TimeStamp $FolderPath does not exist." | out-file -filepath $OutputLoc -append
   return #exit function when folderpath is not found
   }
   elseif((test-path -Path "$FolderPath\$FolderName") -eq $false)
   {
   #create new folder if doesn't exist
   new-item -path $FolderPath -name $FolderName -itemtype $FolderType -force -ErrorAction silentlycontinue
   "$TimeStamp Created registry folder named $FolderName." | out-file -filepath $OutputLoc -append
   return
   }
   }
catch{
     $_.Exception.Message -like "* already exists*" #error ahhh.....
     }
finally{
       #yay
       "$TimeStamp Folder name $FolderName under $FolderPath is already existed." | out-file -filepath $OutputLoc -append
    }
   
}


#Add stratanetwork.com to compatibility view
$ieversion = (get-item -path 'HKLM:\Software\Microsoft\Internet Explorer').getvalue('svcVersion') #IE 11 version reg path
$ie11compviewpath = 'HKCU:\Software\Microsoft\Internet Explorer\BrowserEmulation\ClearableListData' #IE 11 compatiblity view settings path
$ie11key = 'userfilter' #registry key name
$ie11value = (get-item -path 'HKCU:\Software\Microsoft\Internet Explorer\BrowserEmulation\ClearableListData').getvalue('userfilter')
if ($ieversion -like "11*") #check for existing IE version
{
    #Verify stratanetwork.com is on the IE 11 Compatibility View Settings
    
    #check for existing registry value
    if((test-RegistryKey -path $ie11compviewpath -value $ie11key) -eq $false) 
    {
    "$TimeStamp stratanetwork.com needs to be added to the IE 11 Compatibility View Settings." | out-file -filepath $OutputLoc -append
      #stop if returns false
    }
    #check if stratanetwork.com existed as a substring
    elseif(([Text.Encoding]::unicode.getstring($ie11value) -like "*stratanetwork.com*") -ne $false)
    {
    #website existed
    "$TimeStamp stratanetwork.com is already added to the IE 11 Compatibility View Settings." | out-file -filepath $OutputLoc -append
      #exit when the website is confirmed on the list
    }
    else
    {
    #website missing
    "$TimeStamp stratanetwork.com needs to be added to the IE 11 Compatibility View Settings." | out-file -filepath $OutputLoc -append
    
    }
}
else #for all pre-IE 11 versions
{
#go baby go
set-registryvalue -RegPath 'HKLM:\SOFTWARE\Wow6432Node\Policies\Microsoft\Internet Explorer\BrowserEmulation\PolicyList' -regname '*.stratanetwork.com' -regvalue '*.stratanetwork.com' -regtype 'string'
}

#Allow auto prompt
set-registryvalue -RegPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2' -RegName '2200' -RegValue 0 -RegType 'DWORD'
#Add the trusted sites for the current user
set-foldervalue -folderpath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains' -foldername 'stratanetwork.com' -foldertype 'folder'
set-registryvalue -regpath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\stratanetwork.com' -regname 'https' -regvalue 2 -regtype 'dword'
#Add stratanetwork.com to pop-up blocker exception
set-registryvalue -regpath 'HKCU:\Software\Microsoft\Internet Explorer\New Windows\Allow' -regname '*.stratanetwork.com' -regvalue (0X00, 0x00) -type 'binary'
#########################Excel Scripts######################################################
#Enable VBA on Excel
set-registryvalue -regpath 'HKCU:\software\microsoft\office\*\excel\security' -regname 'AccessVBOM' -regvalue 1 -regtype 'dword'
#Add User My Documents location to Trusted Locations
set-foldervalue -folderpath 'HKCU:\software\microsoft\office\*\excel\security\trusted locations' -foldername 'Location99' -foldertype 'folder'
set-registryvalue -regpath 'HKCU:\software\microsoft\office\*\excel\security\trusted locations\Location99' -regname 'Path' -regvalue '$env:USERPROFILE\documents' -regtype 'string'
set-registryvalue -regpath 'HKCU:\software\microsoft\office\*\excel\security\trusted locations\Location99' -regname 'AllowSubfolders' -regvalue 1 -regtype 'dword'
set-registryvalue -regpath 'HKCU:\software\microsoft\office\*\excel\security\trusted locations\Location99' -regname 'Date' -regvalue (get-date).tostring() -regtype 'string'
set-registryvalue -regpath 'HKCU:\software\microsoft\office\*\excel\security\trusted locations\Location99' -regname 'Description' -regvalue 'StrataJazz' -regtype 'string'
set-registryvalue -regpath 'HKCU:\software\microsoft\office\*\excel\security\trusted locations' -regname 'AllowNetworkLocations' -regvalue 1 -regtype 'dword'
#########################Excel Scripts End######################################################


# cross the finish line and show the MsgBox:
$result = [System.Windows.Forms.MessageBox]::Show('Process completed. Please contact support@stratadecision.com for additional assistance.', 'StrataJazz Desktop Requirements Checker', 'OK', 'Information')