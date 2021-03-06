##############################################################
# Detta script skapar delade brevådor i MARTINSERVERA 
#
#    
#
# Created by Jan Lönnman
#  amended and changed: Hans K Hård, 2016-06-03
#
# Version 1.0		2016-06-03
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
Scriptet stödjer INTE <CommonParameters>. 

.PARAMETER Kontonamn
Namn på den nya delade brevlådan som skall skapas. 

.Inputs 
Input skall vara det önskade "DisplayName" som den delade brevlådan skall ha.

.Outputs
Resultatet är en delad brevlåda och en korresponderande SäKerhetsgrupp som har rättigheterna "Full Access" och "Send As".

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

Import-Module ActiveDirectory
. '\\sthdcsrvb174.martinservera.net\Script$\_lib\logfunctions.ps1'
. '\\sthdcsrvb174.martinservera.net\Script$\_lib\connect-exchange.ps1'
connect
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
  if ($_ -eq 'Å' ) { $_ = 'A' }
  if ($_ -eq 'Ä' ) { $_ = 'A' }
  if ($_ -eq 'Ö' ) { $_ = 'O' }
  $tmpStr += $_
 }
 $tmpstr
}
##############################################################


$scriptFileName = ($MyInvocation.MyCommand.Name).split(".")[0]
$logFilePath = "\\sthdcsrvb174.martinservera.net\script$\_log\"
openLogFile "$logFilePath$(($MyInvocation.MyCommand.name).split('.')[0])-$(get-date -uformat %D)-$env:USERNAME.log"
$alreadyExist = $False

#Check parameter "Kontonamn" for sanity                     #TBC 
if ($Kontonamn -notlike "GEM *")
{
 LogLineWithColour -sLine "Namnet $Kontonamn verkar inte vara korrekt!" -sColour "Red"
 Continue
}

LogLineWithColour -sLine "Nu bearbetas brevådan $($Kontonamn)." -sColour "Green"

# Create Alias, UPN, SamAccountName, Password
$Alias = ReplaceSpecialChars($Kontonamn)
$UPN = $Alias + "@martinservera.net" 
$Index = [Math]::Min(20, $Alias.Length)                # $Index must be whitin the string
$Sam = $Alias.Substring(0,$Index) 
$pass = convertto-securestring -string "P@ssw0rd" -asplaintext -force

# Create MailBox 
$alreadyExist = $False 
if ([bool](Get-Mailbox -Identity ([string]$Kontonamn) -ErrorAction SilentlyContinue -DomainController $msDC))
{
 LogLineWithColour ("Brevlådan '$Kontonamn' finns redan!", "Yellow")
 $alreadyExist = $True 
}
If (-not $alreadyExist)
{
 New-Mailbox -Name $Kontonamn -Alias $Alias -OrganizationalUnit $msOU -UserPrincipalName $UPN -SamAccountName $Sam -FirstName '' -Initials '' -LastName '' -Password $pass -ResetPasswordOnNextLogon $false -DomainController $msDC >nul 2>&1 
 Set-Mailbox -Identity ([string]$Kontonamn) -Type shared  -DomainController $msDC  | Out-Null
 LogLineWithColour -sLine "Brevlådan $($Kontonamn) har skapats." -sColour "Green"
}

# Is MailBox of type SharedMailbox?
$MailBox = Get-Mailbox -Identity ([string]$Kontonamn) -DomainController $msDC 
If ($MailBox.RecipientTypeDetails -eq "SharedMailbox")
{ 
   # SharedMailbox 
   # Create Security Group 
   $alreadyExist = $False 
   $GroupName = "SÄK " + $Kontonamn 
   if ([bool](Get-ADGroup -Identity ([string]$GroupName) -Server $msDC -ErrorAction SilentlyContinue))
   {
    LogLineWithColour -sLine "Säkerhetsgruppen '$GroupName' finns redan!" -sColour "Yellow"
    $alreadyExist = $True
   }
   If (-not $alreadyExist)
   {
    New-ADGroup -Name $GroupName -GroupCategory Security -GroupScope Global -Path $msPath -Server $msDC  | Out-Null
    LogLineWithColour -sLine "Säkerhetsgruppen $($GroupName) har skapats." -sColour "Green"
   }

   # Set permissions 
   Add-MailboxPermission -Identity ([string]$Kontonamn) -User $GroupName -AccessRights:FullAccess -InheritanceType:All  -DomainController $msDC | Out-Null
   Add-ADPermission -Identity $MailBox.Name -User $GroupName -AccessRights ExtendedRight -ExtendedRights "Send As"  -DomainController $msDC | Out-Null
   LogLineWithColour -sLine "Rättigheter har nu satts på brevlådan $($Kontonamn)." -sColour "Green"
}
else 
{
   LogLineWithColour -sLine "Brevlådan '$Kontonamn' är inte av typen 'SharedMailbox', utan av typen $($MailBox.RecipientTypeDetails)." -sColour "Yellow"
   LogLineWithColour -sLine "därför har ingen säkerhetsgrupp skapats eller några rättigheter satts." -sColour "Yellow"
}  


Write-Host "Kvar att göra är:" -Foreground Red
Write-Host "   I förekommande fall ändra och/eller lägga till e-postadresser på brevlådan." -Foreground Yellow
Write-Host "   Addera användare till säkerhetsgruppen." -Foreground Yellow
