##Quelle: https://github.com/Darkseal/RunningLow

	### Aufgabenplanung
## Trigger: On a schedule "Daily" "8 Uhr"
## Action: Starte ein Programm
## Script: powershell
## Parameter: -executionpolicy bypass -File dateiname-des-scripts.ps1
## Start in: C:\Pfad-zum-Script\

### - - - - - START DYNAMISCHER BEREICH - - - - - 

	### zu pruefende Laufwerke: Setze $null um alle lokalen Laufwerke zu pruefen (ausser netzlaufwerke)
#$drives = @("C","D","E");
$drives = $null;

## Der Grenzwert, ab wann eine Warnung versendet wird
$minSize = 4GB;


	### SMTP Konfiguration
$email_smtp_host = "smtp.yourdomain.com";
$email_smtp_port = 25;
$email_smtp_SSL = 0;
$email_from_address = "speicherplatzcheck@domain.de";

#$email_username = "username@yourdomain.com";
#$email_password = "yourpassword";

##Empfaenger E-Mail, mehrere Empfaenger Trennung mit Komma
$email_to_addressArray = @("edv@domain.local");

### - - - - - ENDE DYNAMISCHER BEREICH - - - - - 

	### Umrechnungen Bytes ind Gigabytes
$Faktor = 1073741824
$Kommastellen = 0
$minSizeGB = [Math]::Round([float]$minSize / $Faktor, $Kommastellen);
$FreeGB = [Math]::Round([float]$disk.Free / $Faktor, $Kommastellen);
$UsedGB = [Math]::Round([float]$disk.Used / $Faktor, $Kommastellen);

	### Berechnung und Ausfuehrung
if ($drives -eq $null -Or $drives -lt 1) {
    $localVolumes = Get-WMIObject win32_volume;
    $drives = @();
    foreach ($vol in $localVolumes) {
        if ($vol.DriveType -eq 3 -And $vol.DriveLetter -ne $null ) {
            $drives += $vol.DriveLetter[0];
        }
    }
}
foreach ($d in $drives) {
    Write-Host ("`r`n");
    Write-Host ("Prüfe Laufwerk " + $d + " ...");
    $disk = Get-PSDrive $d;
	###HIER AUFRUF für MAIL
    if ($disk.Free -lt $minSize) {
        
		Write-Host ("Laufwerk " + $d + " hat weniger als " + $minSizeGB `
            + " GB freien Speicherplatz (" + $FreeGB + " GB). `n" `
			+ "Sende E-Mail an: " + $email_to_addressArray );
        
        $message = new-object Net.Mail.MailMessage;
        $message.From = $email_from_address;
        foreach ($to in $email_to_addressArray) {
            $message.To.Add($to);
        }
        $message.Subject =  ("[RunningLow] WARNUNG: " + $env:computername + " Laufwerk " + $d);
        $message.Subject += (" hat weniger als " + $minSizeGB + " GB frei ");
        $message.Subject += ("(" + $FreeGB + ") GB");
        $message.Body =     "Hallo, `r`n`r`n";
        $message.Body +=    "dies ist eine automatische Nachricht ";
        $message.Body +=    "gesendet von einem Powershell-Script ";
        $message.Body +=    ("um zu informieren, dass " + $env:computername + " Laufwerk " + $d + " ");
        $message.Body +=    "nur noch wenig freien Speicherplatz hat. `r`n`r`n";
        $message.Body +=    "--------------------------------------------------------------";
        $message.Body +=    "`r`n";
        $message.Body +=    ("Machine HostName: " + $env:computername + " `r`n");
        $message.Body +=    "Server IP Addresse(n): ";
        $ipAddresses = Get-NetIPAddress -AddressFamily IPv4;
        foreach ($ip in $ipAddresses) {
            if ($ip.IPAddress -like "127.0.0.1") {
                continue;
            }
            $message.Body += ($ip.IPAddress + " ");
        }
        $message.Body +=    "`r`n";
        $message.Body +=    ("Verwendeter Speicher von Laufwerk " + $d + ": " + $UsedGB + " GB. `r`n");
        $message.Body +=    ("Freier Speicher von Laufwerk " + $d + ": " + $FreeGB + " GB. `r`n");
        $message.Body +=    "--------------------------------------------------------------";
        $message.Body +=    "`r`n`r`n";
        $message.Body +=    "Diese Warnung wird gesendet, wenn der Speicher weniger ";
        $message.Body +=    ("als " + $minSizeGB + " GB `r`n`r`n");
        $message.Body +=    "Alles Gute. `r`n`r`n";
        $message.Body +=    "-- `r`n";
        $message.Body +=    "RunningLow`r`n";

        $smtp = new-object Net.Mail.SmtpClient($email_smtp_host, $email_smtp_port);
        #$smtp.EnableSSL = $email_smtp_SSL;
        #$smtp.Credentials = New-Object System.Net.NetworkCredential($email_username, $email_password);
        $smtp.send($message);
        $message.Dispose();F
        write-host "... E-Mail gesendet!" ; 
    }
    else {
        Write-Host ("Auf Laufwerk " + $d + " sind " + $FreeGB + " GB von "  + $UsedGB + " GB frei: Keine Aktion noetig.");
    }
}
