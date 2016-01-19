Function update-registryvalue{
   Param
   (
   $RegPath,
   $RegName,
   $RegValue,
   $RegType
   
    )
try{   
    #check registry key
   if(Test-Path -path $RegPath)
   {
   #create a new registry key if it doesn't exist
   new-itemproperty -path $RegPath -name $RegName -value $RegValue -type $RegType
   Write-host "Created registry key $RegName." 
   }
   #check registry value
   elseif((get-itemproperty -path $RegPath).$RegName -ne $RegValue)
   {
   #update registry value if it doesn't match
   set-itemproperty -path $RegPath -name $RegName -value $RegValue -type $RegType
   Write-host "Updated registry key $RegName value to $RegValue." 
   }
   }
catch{
     $_.Exception.Message -like "* already exists*"
     "Registry value for key $RegName is already existed." | out-file c:\logs\log.txt -append
     break
   }
finally{
      Get-Date | Add-Content log.txt
      "Registry key $RegName under $RegPath has been updated with $RegValue." | out-file c:\logs\log.txt -append
    }
   
}


#update-value -FolderPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains' -FolderName 'stratanetwork.com' -FolderType 'folder' -RegName 'https' -RegValue 2 -RegType 'Dword'
update-registryvalue -RegPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2' -regname '2200' -regvalue 0 -RegType 'DWORD'
#(get-item 'HKCU:\Software\Microsoft\Internet Explorer\New Windows\Allow').getvalue('*.stratanetwork.com')
#(Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2').2200 -eq 0
