#------------------------------------------------------------------------
# Start-Sleep 5
param(
  [Parameter(Mandatory = $false)]
  [boolean]$Anteprima = $false
)
$Global:Anteprima = $Anteprima
#------------------------------------------------------------------------
$global:excludeUsers = @()
$global:esclusifileName = 'UtentiEsclusi.txt'
$global:scriptPath = $MyInvocation.MyCommand.Path
#------------------------------------------------------------------------

function CaricaEsclusi {
  $global:esclusifileName = 'UtentiEsclusi.txt'
  $currentPath = Split-Path $global:scriptPath
  $filePath = Join-Path -Path $currentPath -ChildPath $global:esclusifileName 
  if (Test-Path $filePath) {
    # Se il file esiste, legge il contenuto nel array
    $global:excludeUsers = Get-Content $filePath
  }
}

function Get-UnixTimestamp {
  [CmdletBinding()]
  param ()
  $unixEpoch = New-Object System.DateTime 1970, 1, 1, 0, 0, 0, 0, 'Utc'
  $dateTime = Get-Date
  $timeSpan = $dateTime - $unixEpoch
  $unixTime = [Math]::Round($timeSpan.TotalSeconds)
  return $unixTime
}

function ConvertToHHMMSS {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [int]$unixTimestamp
  )
  $dateTime = [System.DateTimeOffset]::FromUnixTimeSeconds($unixTimestamp)
  $timeString = $dateTime.ToString("HH:mm:ss")
  return $timeString
}
  
function PremereUnTastoPerContinuare {
  Write-Host "Premere un tasto per continuare..."
  $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function ControllaAdmin {
  $windowsID = [System.Security.Principal.WindowsIdentity]::GetCurrent()
  $windowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($windowsID)
    
  $adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator
  
  if (-not $windowsPrincipal.IsInRole($adminRole)) {
    Write-Host "Questo script deve essere eseguito come amministratore!" -ForegroundColor Red
    PremereUnTastoPerContinuare
    Exit 
  }
}

#------------------------------------------------------------------------
ControllaAdmin
#------------------------------------------------------------------------
CaricaEsclusi
#------------------------------------------------------------------------

$CurrentUserSID = (New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())).Identity.User.Value

$filteredUserProfiles = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.Loaded -eq $false -and $_.Special -eq $false -and $_.SID -ne $CurrentUserSID -and $global:excludeUsers -notcontains $_.LocalPath.split('\')[-1] }

# $profileCount = $filteredUserProfiles.Count
$profileCount = 0
foreach ($UserProfile in $filteredUserProfiles) {
  $profileCount++
}

if ( $profileCount -le 0  ) {
  Write-Output "Non ci sono profili idonei per la cancellazione"
  PremereUnTastoPerContinuare
  Exit
}
if ( $Anteprima ) {
  Write-Output ""
  Write-Output "------- MODALITA' ANTEPRIMA, I DATI NON VERRANNO CANCELLATI -------"
  Write-Output ""
}

$risposta = Read-Host "Attenzione, si sta per eliminate i dati di ${profileCount} profilo/i utente, proseguire? (S|[N])"
if ($risposta -ne "S") {
  Write-Output "Operazione interrotta"
  PremereUnTastoPerContinuare
  Exit
}
    
$elementiCancelati = 0
$errori = 0
$counter = 0
$totalSeconds = 0;
$toDo = $profileCount;
$tempoStrascorsoLabel = ""
$tempoStimatoLabel = ""
    
Write-Output "--------------------------------------------------------"

foreach ($UserProfile in $filteredUserProfiles) {
      
  $sid = New-Object System.Security.Principal.SecurityIdentifier($UserProfile.SID)
  try {
    $user = $sid.Translate([System.Security.Principal.NTAccount])
  }
  catch {
    $user = "undefined"
  }

  $counter++
  Write-Host
  Write-Output "Eliminazione del profilo $counter $user ($toDo/$profileCount) ${tempoStrascorsoLabel} ${tempoStimatoLabel}"
  $toDo--
      
  $start = Get-UnixTimestamp
      
  try {
    if ( -not $Anteprima ) {
      $UserProfileToDelete = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.SID -eq $UserProfile.SID }
      $UserProfileToDelete | Remove-WmiObject
    }
    $elementiCancelati++
  }
  catch {
    Write-Host
    Write-Host "Errore: " $_.Exception.Message
    $errori++
    if ( -not $Anteprima ) {
      $userDirectoryPath = $UserProfileToDelete.LocalPath
      if ((Test-Path -Path $userDirectoryPath) -and $userDirectoryPath.ToLower().StartsWith("c:\users\") -and $userDirectoryPath.Length -gt 9 ) {
        Remove-Item -Path $userDirectoryPath -Recurse -Force -ErrorAction SilentlyContinue
      }
    }
  }

  $stop = Get-UnixTimestamp
  
  $d = $stop - $start;
  $totalSeconds = $totalSeconds + $d

  $tempoStimato = ($totalSeconds / $counter) * $toDo;

  $tempoStrascorsoHHMMSS = ConvertToHHMMSS -unixTimestamp $totalSeconds
  $tempoStimatoHHMMSS = ConvertToHHMMSS -unixTimestamp $tempoStimato
  $tempoStrascorsoLabel = " Tempo trascorso: $tempoStrascorsoHHMMSS"
  $tempoStimatoLabel = " Tempo stimato: $tempoStimatoHHMMSS"
  
}

$tempoStrascorsoHHMMSS = ConvertToHHMMSS -unixTimestamp $totalSeconds
Write-Output "--------------------------------------------------------"
Write-Output "Fine pulizia profili"
Write-Output "Profili cancellati $elementiCancelati"
Write-Output "Tempo trascorso $tempoStrascorsoHHMMSS"
Write-Output "Errori $errori"
PremereUnTastoPerContinuare