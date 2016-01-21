# Add the required .NET assembly:
Add-Type -AssemblyName System.Windows.Forms

$TimeStamp = get-date -format g
$OutputLoc = '.\Documents\Strata Decision Technology\StrataJazzScriptLog.txt'

#create a function to check existing registry value
function test-RegistryValue {

param (
 [parameter(Mandatory=$true)]
 [ValidateNotNullOrEmpty()]$Path,
 [parameter(Mandatory=$true)]
 [ValidateNotNullOrEmpty()]$Value
)

 

try {
    Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Value -force -ErrorAction Stop | Out-Null

    return $true
 }
catch {
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
   if(Test-Path -path $RegPath -eq $false)
   {
   "$TimeStamp $RegPath does not exist." | out-file -filepath $OutputLoc -append
   break
   #exit function since registry path not found
   }
   elseif(test-registryvalue -path $RegPath -value $RegName -eq $false)
   {
   #create a new registry key if it doesn't exist
   new-itemproperty -path $RegPath -name $RegName -value $RegValue -type $RegType -force
   "$TimeStamp Created registry key $RegName." | out-file -filepath $OutputLoc -append
   }
   #check registry value
   elseif((get-itemproperty -path $RegPath).$RegName -ne $RegValue)
   {
   #update registry value if it doesn't match
   set-itemproperty -path $RegPath -name $RegName -value $RegValue -type $RegType -force -erroraction stop
   "$TimeStamp Updated registry key $RegName value to $RegValue." | out-file -filepath $OutputLoc -append 
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
   if(Test-Path -path $FolderPath -eq $false)
   {
   "$TimeStamp $FolderPath does not exist." | out-file -filepath $OutputLoc -append
   break #exit function when folderpath is not found
   }
   elseif(test-path -Path "$FolderPath\$FolderName" -eq $false)
   {
   #create new folder if doesn't exist
   new-item -path $FolderPath -name $FolderName -itemtype $FolderType -force -ErrorAction stop
   "$TimeStamp Created registry key $RegName." | out-file -filepath $OutputLoc -append
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
#Allow auto prompt
#set-registryvalue -RegPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2' -RegName '2200' -RegValue 0 -RegType 'DWORD'
#Add stratanetwork.com to compatibility view
#Add the trusted sites for the current user
#set-foldervalue -folderpath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains' -foldername 'stratanetwork.com' -foldertype 'folder'
#set-registryvalue -regpath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\stratanetwork.com' -regname 'https' -regvalue 2 -regtype 'dword'
set-registryvalue -regpath 'HKCU:\Software\Microsoft\Internet Explorer\New Windows\Allow' -regname '*.stratanetwork.com' -regvalue (0X00, 0x00) -type binary
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
$ieversion = (get-item -path 'HKLM:\Software\Microsoft\Internet Explorer').getvalue('svcVersion') #IE 11 version reg path
$ie11compviewpath = 'HKCU:\Software\Microsoft\Internet Explorer\BrowserEmulation\ClearableListData' #IE 11 compatiblity view settings path
$ie11value = 'userfilter' #registry key name
if ($ieversion -like "11*") #check for existing IE version
{
    #Verify stratanetwork.com is on the IE 11 Compatibility View Settings
    $currentlist = (get-item -path $ie11compviewpath).getvalue($ie11value)
    #check for existing registry value
    if(((test-RegistryValue -path $ie11compviewpath -value $ie11value) -eq $false) -or ($currentlist -eq $null))
    {
    "$TimeStamp stratanetwork.com needs to be added to the IE 11 Compatibility View Settings." | out-file -filepath $OutputLoc -append
    }
    #check if stratanetwork.com existed as a substring
    #stupid machine wont generate registry key userfilter - errrr
    elseif ([Text.Encoding]::unicode.getstring($value) -like "*stratanetwork.com*" -eq $true)
    {
    #website existed
    "$TimeStamp stratanetwork.com is already added to the IE 11 Compatibility View Settings." | out-file -filepath $OutputLoc -append
    }
}
else #for all pre-IE 11 versions
{
#go baby go
set-registryvalue -RegPath 'HKLM:\SOFTWARE\Wow6432Node\Policies\Microsoft\Internet Explorer\BrowserEmulation\PolicyList' -regname '*.stratanetwork.com' -regvalue '*.stratanetwork.com' -regtype 'string'
}

# cross the finish line and show the MsgBox:
$result = [System.Windows.Forms.MessageBox]::Show('Setup completed. Please contact support@stratadecision.com for additional assistance.', 'StrataJazz Desktop Requirements Checker', 'OK', 'Information')
