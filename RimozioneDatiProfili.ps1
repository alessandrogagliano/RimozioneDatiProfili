#------------------------------------------------------------------------
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
$global:logExpiration = 7 # giorni
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

function Convert-To-HHMMSS {
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
    Write-Host ""
    Write-Host $msg
    Write-Host ""
    Write-Log $msg
    Write-Log "[ERRORE]: $_"
    PremereUnTasto
    Return
  }
  
  $profileCount = 0
  foreach ($UserProfile in $filteredUserProfiles) {
    $profileCount++
  }
  
  if ( $profileCount -le 0  ) {
    Write-Host ""
    Write-Host "Non ci sono profili idonei per la cancellazione"
    Write-Host ""
    PremereUnTasto
    Return
  }
  
  Write-Host ""
  Write-Host "ATTENZIONE, i dati di ${profileCount} profilo/i utente verranno cancellati definitivamente"
  Write-Host "L'OPERAZIONE NON POTRA' ESSERE ANNULLATA"
  Write-Host "Premere ESC per interrompere il processo di cancellazione alla fine dell'operazione in corso"
  Write-Host ""
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
      
  Write-Host "--------------------------------------------------------"
  
  foreach ($UserProfile in $filteredUserProfiles) {

    if ([System.Console]::KeyAvailable) {
      $keyInfo = [System.Console]::ReadKey($true)
      # Controlla se l'utente ha premuto il tasto ESC
      if ($keyInfo.Key -eq "Escape") { 
        # Termina il loop quando viene premuto il tasto ESC.
        PremereUnTasto "E' stato premuto il tasto ESC, operazione interrotta. Premere un tasto per continuare"
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
    Write-Host "Eliminazione del profilo $counter $user ($toDo/$profileCount) ${tempoStrascorsoLabel} ${tempoStimatoLabel}"
    Write-Log "Eliminazione dati utente ${user} ($UserHomePath)"
    $toDo--
    
    $start = Get-UnixTimestamp
    
    # Primo tentaivo di eliminazione
    try {
      $UserProfile | Remove-WmiObject
      $elementiCancelati++
    }
    catch {
      Write-Host "Errore: $_"
      Write-Log "[ERRORE]: $_"
    }

    $userDirectoryPath = $UserProfile.LocalPath
    $RegProfilePath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\"
    $RegProfileKey = $RegProfilePath + $UserProfile.SID

    # Verifica l'effettiva eliminazione della chiave di registro
    if ((Test-Path -Path $RegProfileKey) -and ($RegProfilePath -ne $RegProfileKey)) {
      try {
        $msg = "Eliminazione chiave di registro $RegProfileKey"
        Write-Host $msg
        Write-Log $msg
        Remove-Item -Path $RegProfileKey -Recurse
      }
      catch {
        Write-Host "Errore: $_"
        Write-Log "[ERRORE]: $_"
      }
    }

    # Verifica l'effettiva eliminazione della directory
    if ((Test-Path -Path $userDirectoryPath) -and $userDirectoryPath.ToLower().StartsWith("c:\users\") -and $userDirectoryPath.Length -gt 9 ) {
      try {
        $msg = "Eliminazione directory $userDirectoryPath"
        Write-Host $msg
        Write-Log $msg
        Remove-Item -Path $userDirectoryPath -Recurse -Force -ErrorAction SilentlyContinue
      }
      catch {
        Write-Host "Errore: $_" 
        Write-Log "[ERRORE]: $_"
      }
    }

    # Verifica l'effettiva eliminazione della chiave di registro e della directory
    if ((Test-Path -Path $RegProfileKey) -or (Test-Path -Path $userDirectoryPath)) {
      $errori++
      $msg = "Errore imprevisto, impossibile procedere con la cancellazione dei dati dell'utente"
      Write-Host $msg
      Write-Log "[ERRORE]: $msg"
    }
  
    $stop = Get-UnixTimestamp
    
    $d = $stop - $start;
    $totalSeconds = $totalSeconds + $d
    $tempoStimato = ($totalSeconds / $counter) * $toDo;
  
    $tempoStrascorsoHHMMSS = Convert-To-HHMMSS -unixTimestamp $totalSeconds
    $tempoStimatoHHMMSS = Convert-To-HHMMSS -unixTimestamp $tempoStimato
    $tempoStrascorsoLabel = " Tempo trascorso: $tempoStrascorsoHHMMSS"
    $tempoStimatoLabel = " Tempo stimato: $tempoStimatoHHMMSS"

  }
  
  $tempoStrascorsoHHMMSS = Convert-To-HHMMSS -unixTimestamp $totalSeconds
  Write-Host ""
  Write-Host "--------------------------------------------------------"
  Write-Host "Fine pulizia profili"
  Write-Host "Profili cancellati $elementiCancelati"
  Write-Host "Tempo trascorso $tempoStrascorsoHHMMSS"
  Write-Host "Errori $errori"
  Write-Host "--------------------------------------------------------"
  Write-Host ""
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