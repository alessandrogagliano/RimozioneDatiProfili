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
$global:scriptDirectory = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$today = Get-Date -Format "yyyy-MM-dd"
$global:scriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$global:logFileExt = "log"
$global:logFileName = $global:scriptName + "-" + $today + "." + $global:logFileExt
$global:logFilePath = Join-Path -Path $global:currentPath -ChildPath $global:logFileName
$global:logExpiration = 1 # giorni
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

function Remove-OldLog {
  $AllFiles = Get-ChildItem -Path $global:scriptDirectory -File
  foreach ($file in $AllFiles) {
    # Estrai la data dal nome del file se esiste nel formato "yyyy-MM-dd"
    if ($file.BaseName -match "\d{4}-\d{2}-\d{2}") {
      $dateStr = $Matches[0]
      $fileDate = [datetime]::ParseExact($dateStr, "yyyy-MM-dd", $null)
      # Se il file è più vecchio di 7 giorni, rimuovilo
      if ((Get-Date).AddDays(-$global:logExpiration) -gt $fileDate) {
        # Write-Host "Rimuovo il file" $file.FullName
        Remove-Item $file.FullName
      }
    }
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
  
function PremereUnTasto {
  param (
    [Parameter(Mandatory = $false)]
    [string]
    # $Message = "Premere un tasto per continuare..."
    $Message = "Premere un tasto per terminare o chiudere la finestra..."
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
    PremereUnTasto
    Return
  }
  
  # $profileCount = $filteredUserProfiles.Count
  $profileCount = 0
  foreach ($UserProfile in $filteredUserProfiles) {
    $profileCount++
  }
  
  if ( $profileCount -le 0  ) {
    Write-Output "Non ci sono profili idonei per la cancellazione"
    PremereUnTasto
    Return
  }
  if ( $Anteprima ) {
    Write-Output ""
    Write-Output "------- MODALITA' ANTEPRIMA, I DATI NON VERRANNO CANCELLATI -------"
    Write-Output ""
  }
  
  Write-Output ""
  Write-Output "Attenzione, si sta per eliminate i dati di ${profileCount} profilo/i utente"
  Write-Output "(Per interrompere l'operazione premere ESC e attendere la fine del processo in corso)"
  Write-Output ""
  $risposta = Read-Host "Confermare? (S|[N])"
  if ($risposta -ne "S") {
    Write-Log "Operazione interrotta"
    PremereUnTasto "Operazione interrotta, premere un tasto per terminare o chiudere la finestra"
    Return
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

    if ([System.Console]::KeyAvailable) {
      $keyInfo = [System.Console]::ReadKey($true)
  
      # Controlla se l'utente ha premuto il tasto ESC
      if ($keyInfo.Key -eq "Escape") { 
        # Termina il loop quando viene premuto il tasto ESC.
        PremereUnTasto "E' stato premuto il tasto ESC, operazione interrotta. Premere un tasto per terminare o chiudere la finestra"
        break  
      }
    }

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
  Write-Output ""
  Write-Output "--------------------------------------------------------"
  Write-Output "Fine pulizia profili"
  Write-Output "Profili cancellati $elementiCancelati"
  Write-Output "Tempo trascorso $tempoStrascorsoHHMMSS"
  Write-Output "Errori $errori"
  Write-Output ""
  PremereUnTasto
}

#------------------------------------------------------------------------
ControllaAdmin
#------------------------------------------------------------------------
CaricaEsclusi
#------------------------------------------------------------------------
RimozioneDatiProfiliRoutine
#------------------------------------------------------------------------
Remove-OldLog
#------------------------------------------------------------------------
