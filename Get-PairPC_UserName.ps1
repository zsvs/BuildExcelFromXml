chcp 65001
$PCHash = @{} #! HashTable for storing info about PC/User pair

$MainOU = #! Path to OU where computers stored 
$UsersOU = #! Path to OU where users stored 

$Ob = Get-CMDevice -CollectionName #! Name of your Device Collection in SCCM

foreach($Object in $Ob)
  {
    $Name = $Object.Name # Get current SCCM object PC name
    $UserNM = $Object.UserName # Get current SCCM object user name (sAMAccountName attribute)

    $ADComputer = Get-ADComputer -SearchBase $MainOU -Filter {Name -like $Name} -Properties * # Get AD computer by name given from SCCM object
    $ADUser = Get-ADUser -SearchBase $UsersOU  -Filter {SamAccountName -like $UserNM} # Get AD user by sAMAccountName given from SCCM object
   


    $PCHash.Add($Object.Name, ($Object.UserName, $ADUser.Name, $Object.PrimaryUser, $ADComputer.extensionAttribute1, $ADComputer.extensionAttribute2, $ADComputer.'ms-Mcs-AdmPwd'))
  }




Remove-Item -Path "C:\Users\$env:USERNAME\Documents\PC2User.xml"
Export-Clixml -InputObject $PCHash -Path "C:\Users\$env:USERNAME\Documents\PC2User.xml"