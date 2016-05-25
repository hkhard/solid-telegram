# Detta script skapar delade brevådor i MARTINSERVERA 
#
#    
#
# Created by Jan Lönnman 
#
# Version 0.9.5		2012-03-15
#  
##############################################################

<#
.SYNOPSIS
Detta script skapar delade brevlådor i MARTINSERVERA.

.DESCRIPTION
Detta script utför följande: 
   Skapar en delad brevlåda 
   Skapar en säkerhetsgrupp 
   Ger gruppen rättigheterna "Full Access" och "Send As" för brevlådan 


NB!!!
Scriptet stöder INTE <CommonParameters>. 

.PARAMETER Kontonamn
Namn på den nya delade brevlådan som skall skapas. 

.Inputs 
Input skall vara det önskade "DisplayName" som den delade brevlådan skall ha.

.Outputs
Resultatet är en delad brevlåda och en korresponderande SÄKerhetsgrupp som har rättigheterna "Full Access" och "Send As".

.EXAMPLE
PS \> .\SkapaGEMBrevlåda.ps1 "GEM HSTL1 Halmstad@servera.se"
Detta kommando visar ett typiskt användande; det skapa en ny delad brevlåda som i exemplet heter "GEM HSTL1 Halmstad@servera.se" och en tillhörande säkerhetsgrupp som heter "SÄK GEM HSTL1 Halmstad@servera.se".

.EXAMPLE
PS \> Get-Content InputFile.txt | .\SkapaGEMBrevlåda.ps1
Detta kommando skickar innehållet i en textfil vidara till detta script. 

.EXAMPLE
PS \> "GEM HSTL1 Halmstad@servera.se", "GEM NRKL1 VRM DF@martinservera.se" | .\SkapaGEMBrevlåda.ps1
Detta kommando skickar in två namn till detta script. 

#>

[CmdletBinding()]
param(
  [parameter(Position=0,Mandatory=$True,ValueFromPipeline=$TRUE,HelpMessage="Ange namnet på det/de konton som skall bahandlas. Tryck enter när du skrivit in alla.")] [String[]] $Kontonamn
)

##############################################################
BEGIN {
Import-Module ActiveDirectory 
$ErrorActionPreference = "SilentlyContinue"
$msDC = "STHDCSRV169.martinservera.net" 
$msOU = "martinservera.net/Exchangeresurser"
$msPath = "OU=Epost,OU=Rättigheter,OU=Grupper,DC=martinservera,DC=net" 
##############################################################
function ReplaceSpecialChars([string]$str) {
 $str.ToCharArray() | foreach {
  if ($_ -eq ' ' ) { $_ = '' }
  if ($_ -eq ':' ) { $_ = '' }
  if ($_ -eq '.' ) { $_ = '' }
  if ($_ -eq '@' ) { $_ = '' }
  if ($_ -eq 'å' ) { $_ = 'a' }
  if ($_ -eq 'ä' ) { $_ = 'a' }
  if ($_ -eq 'ö' ) { $_ = 'o' }
  if ($_ -eq 'Å' ) { $_ = 'Å' }
  if ($_ -eq 'Ä' ) { $_ = 'Ä' }
  if ($_ -eq 'Ö' ) { $_ = 'Ö' }
  $tmpStr += $_
 }
 $tmpstr
}
##############################################################
}  #End BEGIN

PROCESS {
$alreadyExist = $False

#Check parameter "Kontonamn" for sanity                     #TBC 
if ($Kontonamn -notlike "GEM *") {
   Write-Host "Namnet $Kontonamn verkar inte korrekt!" -Foreground red
   Continue
}

Write-Host "Nu bearbetas brevådan '$Kontonamn'.`n`n" -Foreground green

# Create Alias, UPN, SamAccountName, Password
$Alias = ReplaceSpecialChars($Kontonamn)
$UPN = $Alias + "@martinservera.net" 
$Index = [Math]::Min(20, $Alias.Length)                # $Index must be whitin the string
$Sam = $Alias.Substring(0,$Index) 
$pass = convertto-securestring -string "P@ssw0rd" -asplaintext -force

# Create MailBox 
$alreadyExist = $False 
if ([bool](Get-Mailbox -Identity ([string]$Kontonamn) -ErrorAction SilentlyContinue -DomainController $msDC)) {
   Write-Host "Brevlådan '$Kontonamn' finns redan!`n" -Foreground yellow
   $alreadyExist = $True 
}
If (-not $alreadyExist) {
New-Mailbox -Name $Kontonamn -Alias $Alias -OrganizationalUnit $msOU -UserPrincipalName $UPN -SamAccountName $Sam -FirstName '' -Initials '' -LastName '' -Password $pass -ResetPasswordOnNextLogon $false -DomainController $msDC 
# | Out-Null
Set-Mailbox -Identity ([string]$Kontonamn) -Type shared  -DomainController $msDC  | Out-Null
Write-Host "Brevlådan $Kontonamn' har skapats.`n" -Foreground green
}

# Is MailBox of type SharedMailbox?
$MailBox = Get-Mailbox -Identity ([string]$Kontonamn) -DomainController $msDC 
If ($MailBox.RecipientTypeDetails -eq "SharedMailbox") {   # SharedMailbox 

   # Create Security Group 
   $alreadyExist = $False 
   $GroupName = "SÄK " + $Kontonamn 
   if ([bool](Get-ADGroup -Identity ([string]$GroupName) -Server $msDC -ErrorAction SilentlyContinue)) {
      Write-Host "Säkerhetsgruppen '$GroupName' finns redan!`n" -Foreground yellow
      $alreadyExist = $True
   }
   If (-not $alreadyExist) {
   New-ADGroup -Name $GroupName -GroupCategory Security -GroupScope Global -Path $msPath -Server $msDC  | Out-Null
   Write-Host "Säkerhetsgruppen '$GroupName' har skapats.`n" -Foreground green
   }

   # Set permissions 
   Add-MailboxPermission -Identity ([string]$Kontonamn) -User $GroupName -AccessRights:FullAccess –InheritanceType All  -DomainController $msDC | Out-Null
   Add-ADPermission -Identity $MailBox.Name -User $GroupName -AccessRights ExtendedRight -ExtendedRights "Send As"  -DomainController $msDC | Out-Null
   Write-Host "Rättigheter har nu satts på brevlådan '$Kontonamn'.`n"  -Foreground green
   }
else {
   Write-Host "Brevlådan '$Kontonamn' är inte av typen 'SharedMailbox', utan av typen" $MailBox.RecipientTypeDetails -Foreground yellow
   Write-Host "därför har ingen säkerhetsgrupp skapats eller några rättigheter satts.`n" -Foreground yellow
}  # Is MailBox of type SharedMailbox? 


} #End PROCESS

END {
#Write-Host "`n`nOch nu har brevlåda skapats.`n" -Foreground green
Write-Host "`nKvar att göra är:`n" -Foreground green
Write-Host "   I förekommande fall ändra och/eller lägga till e-postadresser på brevlåda." -Foreground green
Write-Host "   Addera användare till grupp.`n" -Foreground green
}






##############################################################
#Debug 
<# 

Inparameter (namnet på GEM) 
	kontrollera rimligt format på GEM 

Om GEM inte finns som brevlåda av någon typ
	skapa brevlåda av typ Shared
sedan

	Om GEM finns som brevlåda av typ Shared 
		om grupp SÄK GEM inte finns 
			skapa grupp SÄK GEM 
		sätt fulla rättigheter till GEM med gruppen SÄK GEM 

#<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#>  
