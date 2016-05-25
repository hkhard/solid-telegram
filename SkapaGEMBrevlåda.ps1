# Detta script skapar delade brevådor i MARTINSERVERA 
#
#    
#
# Created by Jan Lönnman 
#
# Version 0.9.6		2016-05-25
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
Scriptet st�der INTE <CommonParameters>. 

.PARAMETER Kontonamn
Namn p� den nya delade brevl�dan som skall skapas. 

.Inputs 
Input skall vara det �nskade "DisplayName" som den delade brevl�dan skall ha.

.Outputs
Resultatet �r en delad brevl�da och en korresponderande S�Kerhetsgrupp som har r�ttigheterna "Full Access" och "Send As".

.EXAMPLE
PS \> .\SkapaGEMBrevl�da.ps1 "GEM HSTL1 Halmstad@servera.se"
Detta kommando visar ett typiskt anv�ndande; det skapa en ny delad brevl�da som i exemplet heter "GEM HSTL1 Halmstad@servera.se" och en tillh�rande s�kerhetsgrupp som heter "S�K GEM HSTL1 Halmstad@servera.se".

.EXAMPLE
PS \> Get-Content InputFile.txt | .\SkapaGEMBrevl�da.ps1
Detta kommando skickar inneh�llet i en textfil vidara till detta script. 

.EXAMPLE
PS \> "GEM HSTL1 Halmstad@servera.se", "GEM NRKL1 VRM DF@martinservera.se" | .\SkapaGEMBrevl�da.ps1
Detta kommando skickar in tv� namn till detta script. 

#>

[CmdletBinding()]
param(
  [parameter(Position=0,Mandatory=$True,ValueFromPipeline=$TRUE,HelpMessage="Ange namnet p� det/de konton som skall bahandlas. Tryck enter n�r du skrivit in alla.")] [String[]] $Kontonamn
)

##############################################################
BEGIN {
Import-Module ActiveDirectory 
$ErrorActionPreference = "SilentlyContinue"
$msDC = "STHDCSRV169.martinservera.net" 
$msOU = "martinservera.net/Exchangeresurser"
$msPath = "OU=Epost,OU=R�ttigheter,OU=Grupper,DC=martinservera,DC=net" 
##############################################################
function ReplaceSpecialChars([string]$str) {
 $str.ToCharArray() | foreach {
  if ($_ -eq ' ' ) { $_ = '' }
  if ($_ -eq ':' ) { $_ = '' }
  if ($_ -eq '.' ) { $_ = '' }
  if ($_ -eq '@' ) { $_ = '' }
  if ($_ -eq '�' ) { $_ = 'a' }
  if ($_ -eq '�' ) { $_ = 'a' }
  if ($_ -eq '�' ) { $_ = 'o' }
  if ($_ -eq '�' ) { $_ = '�' }
  if ($_ -eq '�' ) { $_ = '�' }
  if ($_ -eq '�' ) { $_ = '�' }
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

Write-Host "Nu bearbetas brev�dan '$Kontonamn'.`n`n" -Foreground green

# Create Alias, UPN, SamAccountName, Password
$Alias = ReplaceSpecialChars($Kontonamn)
$UPN = $Alias + "@martinservera.net" 
$Index = [Math]::Min(20, $Alias.Length)                # $Index must be whitin the string
$Sam = $Alias.Substring(0,$Index) 
$pass = convertto-securestring -string "P@ssw0rd" -asplaintext -force

# Create MailBox 
$alreadyExist = $False 
if ([bool](Get-Mailbox -Identity ([string]$Kontonamn) -ErrorAction SilentlyContinue -DomainController $msDC)) {
   Write-Host "Brevl�dan '$Kontonamn' finns redan!`n" -Foreground yellow
   $alreadyExist = $True 
}
If (-not $alreadyExist) {
New-Mailbox -Name $Kontonamn -Alias $Alias -OrganizationalUnit $msOU -UserPrincipalName $UPN -SamAccountName $Sam -FirstName '' -Initials '' -LastName '' -Password $pass -ResetPasswordOnNextLogon $false -DomainController $msDC 
# | Out-Null
Set-Mailbox -Identity ([string]$Kontonamn) -Type shared  -DomainController $msDC  | Out-Null
Write-Host "Brevl�dan $Kontonamn' har skapats.`n" -Foreground green
}

# Is MailBox of type SharedMailbox?
$MailBox = Get-Mailbox -Identity ([string]$Kontonamn) -DomainController $msDC 
If ($MailBox.RecipientTypeDetails -eq "SharedMailbox") {   # SharedMailbox 

   # Create Security Group 
   $alreadyExist = $False 
   $GroupName = "S�K " + $Kontonamn 
   if ([bool](Get-ADGroup -Identity ([string]$GroupName) -Server $msDC -ErrorAction SilentlyContinue)) {
      Write-Host "S�kerhetsgruppen '$GroupName' finns redan!`n" -Foreground yellow
      $alreadyExist = $True
   }
   If (-not $alreadyExist) {
   New-ADGroup -Name $GroupName -GroupCategory Security -GroupScope Global -Path $msPath -Server $msDC  | Out-Null
   Write-Host "S�kerhetsgruppen '$GroupName' har skapats.`n" -Foreground green
   }

   # Set permissions 
   Add-MailboxPermission -Identity ([string]$Kontonamn) -User $GroupName -AccessRights:FullAccess �InheritanceType All  -DomainController $msDC | Out-Null
   Add-ADPermission -Identity $MailBox.Name -User $GroupName -AccessRights ExtendedRight -ExtendedRights "Send As"  -DomainController $msDC | Out-Null
   Write-Host "R�ttigheter har nu satts p� brevl�dan '$Kontonamn'.`n"  -Foreground green
   }
else {
   Write-Host "Brevl�dan '$Kontonamn' �r inte av typen 'SharedMailbox', utan av typen" $MailBox.RecipientTypeDetails -Foreground yellow
   Write-Host "d�rf�r har ingen s�kerhetsgrupp skapats eller n�gra r�ttigheter satts.`n" -Foreground yellow
}  # Is MailBox of type SharedMailbox? 


} #End PROCESS

END {
#Write-Host "`n`nOch nu har brevl�da skapats.`n" -Foreground green
Write-Host "`nKvar att g�ra �r:`n" -Foreground green
Write-Host "   I f�rekommande fall �ndra och/eller l�gga till e-postadresser p� brevl�da." -Foreground green
Write-Host "   Addera anv�ndare till grupp.`n" -Foreground green
}






##############################################################
#Debug 
<# 

Inparameter (namnet p� GEM) 
	kontrollera rimligt format p� GEM 

Om GEM inte finns som brevl�da av n�gon typ
	skapa brevl�da av typ Shared
sedan

	Om GEM finns som brevl�da av typ Shared 
		om grupp S�K GEM inte finns 
			skapa grupp S�K GEM 
		s�tt fulla r�ttigheter till GEM med gruppen S�K GEM 

#<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#>  
