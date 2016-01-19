# Add the required .NET assembly:
Add-Type -AssemblyName System.Windows.Forms

$TimeStamp = get-date -format g
$OutputLoc = '.\Documents\Strata Decision Technology\StrataJazzScriptLog.txt'

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
   if(!(Test-Path -path $RegPath))
   {
   "$TimeStamp $RegPath does not exist." | out-file -filepath $OutputLoc -append
  
   }
   elseif(!(test-registryvalue -path $RegPath -value $RegName))
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
     $_.Exception.Message -like "* already exists*"
     }
finally{
      
       "$TimeStamp Registry value $RegValue for registry key named $RegName is already set correctly." | out-file -filepath $OutputLoc -append
    }
   
}

Function set-FolderValue{
   Param
   (
   $FolderPath,
   $FolderName,
   $FolderType
   )
try{   
   #check folder path
   if(!(Test-Path -path $FolderPath))
   {
   "$TimeStamp $FolderPath does not exist." | out-file -filepath $OutputLoc -append
   }
   elseif(!(test-path -Path "$FolderPath\$FolderName"))
   {
   #create new folder if doesn't exist
   new-item -path $FolderPath -name $FolderName -itemtype $FolderType -force -ErrorAction stop
   "$TimeStamp Created registry key $RegName." | out-file -filepath $OutputLoc -append
   }
   }
catch{
     $_.Exception.Message -like "* already exists*"
     }
finally{
      
       "$TimeStamp Folder name $FolderName under $FolderPath is already existed." | out-file -filepath $OutputLoc -append
    }
   
}
#Allow auto prompt
set-registryvalue -RegPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2' -RegName '2200' -RegValue 0 -RegType 'DWORD'
#Add stratanetwork.com to compatibility view
set-registryvalue -RegPath 'HKLM:\SOFTWARE\Wow6432Node\Policies\Microsoft\Internet Explorer\BrowserEmulation\PolicyList' -regname '*.stratanetwork.com' -regvalue '*.stratanetwork.com' -regtype 'string'
#Add the trusted sites for the current user
set-foldervalue -folderpath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains' -foldername 'stratanetwork.com' -foldertype 'folder'
set-registryvalue -regpath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\stratanetwork.com' -regname 'https' -regvalue 2 -regtype 'dword'
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

#Add stratanetwork.com to pop-up blocker exception
$popupPath = 'HKCU:\Software\Microsoft\Internet Explorer\New Windows\Allow'
$popupName = '*.stratanetwork.com'
$popupValue = New-Object Byte[] 2
$popupType = 'binary'
new-itemproperty -path $popupPath -name $popupName -value $popupValue -propertytype $popupType -force
"$TimeStamp $popupname has been added to the Pop-up Blocker exceptions." | out-file -filepath $OutputLoc -append

# show the MsgBox:
$result = [System.Windows.Forms.MessageBox]::Show('Setup completed. Please contact support@stratadecision.com for additional assistance.', 'StrataJazz Desktop Requirements Checker', 'OK', 'Information')


