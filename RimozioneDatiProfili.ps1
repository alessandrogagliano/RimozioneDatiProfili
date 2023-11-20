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
$global:currentPath = Split-Path $MyInvocation.MyCommand.Path
$global:logFileName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name) + ".log"
$global:logFilePath = Join-Path -Path $global:currentPath -ChildPath $global:logFileName
#------------------------------------------------------------------------

function CaricaEsclusi {
  $global:esclusifileName = 'UtentiEsclusi.txt'
  $filePath = Join-Path -Path $global:currentPath -ChildPath $global:esclusifileName 
  if (Test-Path $filePath) {
    # Se il file esiste, legge il contenuto nel array
    $global:excludeUsers = Get-Content $filePath
  }
}

function Write-Log {
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $Message
  )

  # Get the current timestamp
  $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

  # Prepend the timestamp to the log message
  $logEntry = "$timestamp : $Message"

  # Append the log entry to the log file
  Add-Content -Path $global:logFilePath -Value $logEntry
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
  
function PremereUnTasto {
  param (
    [Parameter(Mandatory = $false)]
    [string]
    $Message = "Premere un tasto per continuare..."
  )
  Write-Host $Message
  $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function ControllaAdmin {
  $windowsID = [System.Security.Principal.WindowsIdentity]::GetCurrent()
  $windowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($windowsID)
    
  $adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator
  
  if (-not $windowsPrincipal.IsInRole($adminRole)) {
    Write-Host "Questo script deve essere eseguito come amministratore!" -ForegroundColor Red
    PremereUnTasto "Operazione interrotta, premere un tasto per terminare o chiudere la finestra"
    Exit 
  }
}

function RimozioneDatiProfiliRoutine {
  $CurrentUserSID = (New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())).Identity.User.Value

  try {
    $filteredUserProfiles = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.Loaded -eq $false -and $_.Special -eq $false -and $_.SID -ne $CurrentUserSID -and $global:excludeUsers -notcontains $_.LocalPath.split('\')[-1] }
  }
  catch {
    $msg = "Impossibile procedere, potrebbero esserci problemi con i file di sistema, si consiglia di utilizzare il comando: sfc /scannow"
    Write-Host $msg
    Write-Log $msg
    PremereUnTasto "Premere un tasto per terminare o chiudere la finestra"
    Exit
  }
  
  # $profileCount = $filteredUserProfiles.Count
  $profileCount = 0
  foreach ($UserProfile in $filteredUserProfiles) {
    $profileCount++
  }
  
  if ( $profileCount -le 0  ) {
    Write-Output "Non ci sono profili idonei per la cancellazione"
    PremereUnTasto "Premere un tasto per terminare o chiudere la finestra"
    Exit
  }
  if ( $Anteprima ) {
    Write-Output ""
    Write-Output "------- MODALITA' ANTEPRIMA, I DATI NON VERRANNO CANCELLATI -------"
    Write-Output ""
  }
  
  $risposta = Read-Host "Attenzione, si sta per eliminate i dati di ${profileCount} profilo/i utente, proseguire? (S|[N])"
  if ($risposta -ne "S") {
    # Write-Output "Operazione interrotta"
    PremereUnTasto "Operazione interrotta, premere un tasto per terminare o chiudere la finestra"
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
        
    $UserHomePath = $UserProfile.LocalPath
    $sid = New-Object System.Security.Principal.SecurityIdentifier($UserProfile.SID)
    try {
      $user = $sid.Translate([System.Security.Principal.NTAccount])
    }
    catch {
      # $user = "undefined"
      $user = $UserHomePath
    }
  
    $counter++
    Write-Host
    Write-Output "Eliminazione del profilo $counter $user ($toDo/$profileCount) ${tempoStrascorsoLabel} ${tempoStimatoLabel}"
    $toDo--
        
    $start = Get-UnixTimestamp
        
    try {
      if ( -not $Anteprima ) {
        $UserProfile | Remove-WmiObject
        Write-Log "Eliminazione dati utente ${user} ($UserHomePath)"
      }
      $elementiCancelati++
    }
    catch {
      Write-Host
      Write-Host "Errore: " $_.Exception.Message
      Write-Log $_.Exception.Message
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
  PremereUnTasto "Premere un tasto per terminare o chiudere la finestra"
}
#------------------------------------------------------------------------
ControllaAdmin
#------------------------------------------------------------------------
CaricaEsclusi
#------------------------------------------------------------------------
RimozioneDatiProfiliRoutine
#------------------------------------------------------------------------
