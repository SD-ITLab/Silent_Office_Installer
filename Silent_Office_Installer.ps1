$host.ui.RawUI.WindowTitle = "Unattended Microsoft Office Deployment"
[Console]::WindowWidth=102;
[Console]::Windowheight=48;
[Console]::setBufferSize(102,48) #width,height


$Display = {
Write-Host "    _____  .__                                 _____  __    ________   _____  _____.__               
   /     \ |__| ___________  ____  ___________/ ____\/  |_  \_____  \_/ ____\/ ____\__| ____  ____ 
  /  \ /  \|  |/ ___\_  __ \/  _ \/  ___/  _ \   __\\   __\  /   |   \   __\\   __\|  |/ ___\/ __ \
 /    Y    \  \  \___|  | \(  <_> )___ (  <_> )  |   |  |   /    |    \  |   |  |  |  \  \__\  ___/
 \____|__  /__|\___  >__|   \____/____  >____/|__|   |__|   \_______  /__|   |__|  |__|\___  >___  > 
         \/        \/                 \/                            \/                     \/    \/" -ForegroundColor Cyan           

Write-Host "                 ________                .__                                      __ 
                 \______ \   ____ ______ |  |   ____ ___.__. _____   ____   _____/  |_
                  |    |  \_/ __ \\____ \|  |  /  _ <   |  |/     \_/ __ \ /    \   __\
                  |    |   \  ___/|  |_> >  |_(  <_> )___  |  Y Y  \  ___/|   |  \  |
                 /_______  /\___  >   __/|____/\____// ____|__|_|  /\___  >___|  /__|
   @sd-itlab.de          \/     \/|__|               \/          \/     \/     \/            v2.0.0 " -ForegroundColor Green
}


$installcheck = {
    # Überprüfen, ob Office installiert ist
    if (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -Match "Microsoft Office" }) {
        Write-Host "| Software ist bereits Installiert  " -NoNewline -ForegroundColor Green
        Write-Host "║" -ForegroundColor Yellow
    } else {
        # Pfad zur Office Setup Datei
        $officeSetupPath = ".\_Office-Installer\setup.exe"
        
        if (Test-Path $officeSetupPath -PathType Leaf) {
            # Office Konfiguration Datei
            $configFile = ".\_Office-Installer\configuration.xml"
            
            # Überprüfen, ob Konfigurationsdatei existiert
            if (Test-Path $configFile -PathType Leaf) {
                Start-Process -Wait -FilePath $officeSetupPath -ArgumentList "/configure $configFile" -NoNewWindow -PassThru | Out-Null
                Write-Host "| Erfolgreich!                      " -ForegroundColor Green -NoNewline
                Write-Host "║" -ForegroundColor Yellow
            } else {
                Write-Host "| FEHLER - Konfigurationsdatei nicht gefunden. " -ForegroundColor Red -NoNewline
                Write-Host "║" -ForegroundColor Yellow
            }
        } else {
            Write-Host "| FEHLER - Setup Datei nicht gefunden. " -ForegroundColor Red -NoNewline
            Write-Host "║" -ForegroundColor Yellow
        }
    }
    Start-Sleep 1
}

$MSHS21 = {
    OfficeConf -ProductID "ProPlus2021Retail"
    Write-Host "  ║  " -ForegroundColor Yellow -NoNewline
    Write-Host "Installiere Software: " -ForegroundColor Red -NoNewline
    Write-Host "Office 2021 Home and Student...   " -ForegroundColor Yellow -NoNewline
    Invoke-Command -ScriptBlock $installcheck
}

$MSHB21 = {
    OfficeConf -ProductID "ProPlus2021Retail"  
    Write-Host "  ║  " -ForegroundColor Yellow -NoNewline
    Write-Host "Installiere Software: " -ForegroundColor Red -NoNewline
    Write-Host "Office 2021 Home and Business...    " -ForegroundColor Yellow -NoNewline
    Invoke-Command -ScriptBlock $installcheck
    
}

$MSPRO21 = {
    OfficeConf -ProductID "ProPlus2021Retail"
    Write-Host "  ║  " -ForegroundColor Yellow -NoNewline
    Write-Host "Installiere Software: " -ForegroundColor Red -NoNewline
    Write-Host "Office 2021 Professional...   " -ForegroundColor Yellow -NoNewline
    Invoke-Command -ScriptBlock $installcheck
}

$MSHS24 = {
    OfficeConf -ProductID "Home2024Retail"
    Write-Host "  ║  " -ForegroundColor Yellow -NoNewline
    Write-Host "Installiere Software: " -ForegroundColor Red -NoNewline
    Write-Host "Office 2024 Home and Student...   " -ForegroundColor Yellow -NoNewline
    Invoke-Command -ScriptBlock $installcheck
}

$MSHB24 = {
    OfficeConf -ProductID "HomeBusiness2024Retail"  
    Write-Host "  ║  " -ForegroundColor Yellow -NoNewline
    Write-Host "Installiere Software: " -ForegroundColor Red -NoNewline
    Write-Host "Office 2024 Home and Business...    " -ForegroundColor Yellow -NoNewline
    Invoke-Command -ScriptBlock $installcheck
    
}

$MSPRO24 = {
    OfficeConf -ProductID "ProPlus2024Retail"
    Write-Host "  ║  " -ForegroundColor Yellow -NoNewline
    Write-Host "Installiere Software: " -ForegroundColor Red -NoNewline
    Write-Host "Office 2024 Professional...   " -ForegroundColor Yellow -NoNewline
    Invoke-Command -ScriptBlock $installcheck
}

$MSHS24P = {
    OfficeConf -ProductID "ProPlus2024Retail"
    Write-Host "  ║  " -ForegroundColor Yellow -NoNewline
    Write-Host "Installiere Software: " -ForegroundColor Red -NoNewline
    Write-Host "Office 2024 Home and Student...   " -ForegroundColor Yellow -NoNewline
    Invoke-Command -ScriptBlock $installcheck
}

$MSHB24P = {
    OfficeConf -ProductID "ProPlus2024Retail"  
    Write-Host "  ║  " -ForegroundColor Yellow -NoNewline
    Write-Host "Installiere Software: " -ForegroundColor Red -NoNewline
    Write-Host "Office 2024 Home and Business...    " -ForegroundColor Yellow -NoNewline
    Invoke-Command -ScriptBlock $installcheck
    
}

$MS365S = {
    OfficeConf -ProductID "O365HomePremRetail"
    Write-Host "  ║  " -ForegroundColor Yellow -NoNewline
    Write-Host "Installiere Software: " -ForegroundColor Red -NoNewline
    Write-Host "Office 356 Single...               " -ForegroundColor Yellow -NoNewline
    Invoke-Command -ScriptBlock $installcheck
}

$MS365F = {
    OfficeConf -ProductID "O365HomePremRetail"
    Write-Host "  ║  " -ForegroundColor Yellow -NoNewline
    Write-Host "Installiere Software: " -ForegroundColor Red -NoNewline
    Write-Host "Office 365 Family...               " -ForegroundColor Yellow -NoNewline
    Invoke-Command -ScriptBlock $installcheck
}



function OfficeConf {
    param (
        [string]$ProductID,
        [string[]]$Apps
    )

    # Create configuration.xml dynamically
    $configContent = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="$ProductID">
      <Language ID="de-de" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="0" />
</Configuration>
"@

    $configFile = ".\_Office-Installer\configuration.xml"
    $configContent | Set-Content -Path $configFile
}

function Shortcut {
    param (
        [string[]]$Apps
    )

    # Creating desktop shortcuts
    $startMenuPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
    $desktopPath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), "")
    foreach ($app in $Apps) {
        $shortcut = [System.IO.Path]::Combine($startMenuPath, "$app.lnk")
        $desktopShortcut = [System.IO.Path]::Combine($desktopPath, "$app.lnk")

        if (Test-Path $shortcut) {
            Copy-Item -Path $shortcut -Destination $desktopShortcut -Force
        }
        else {
        }
    }
}

$End = {
  Write-Host "  ║                                                                                               ║" -ForegroundColor Yellow
  Write-Host "  ╠═══════════════════════════════════════════════════════════════════════════════════════════════╣" -ForegroundColor Yellow
  Write-Host "  ║                           " -NoNewline -ForegroundColor Yellow
  Write-Host "Der Installationsprozess war Erfolgreich!" -NoNewline -ForegroundColor Green
  Write-Host "                           ║" -ForegroundColor Yellow
  Write-Host "  ╚═══════════════════════════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
  Start-Sleep 5
  }

function menu {

  Invoke-Command -ScriptBlock $Display
  Write-Host " ═════════════╦═══════════════════════════════════════════════════════════════════════╦═════════════" -ForegroundColor Yellow
  Write-Host "              ╠════════════════════════════" -NoNewline -ForegroundColor Yellow
  Write-Host "  Office 2021  " -ForegroundColor White -NoNewline
  Write-Host "════════════════════════════╣              " -ForegroundColor Yellow
  Write-Host "              ║                                                                       ║" -ForegroundColor Yellow
  Write-Host "              ║" -ForegroundColor Yellow -NoNewLine
  Write-Host "    [1]: Microsoft Office 2021 Home and Student (Pro Plus Image)       " -ForegroundColor Cyan -NoNewLine
  Write-Host "║" -ForegroundColor Yellow
  Write-Host "              ║                                                                       ║" -ForegroundColor Yellow
  Write-Host "              ║" -ForegroundColor Yellow -NoNewLine
  Write-Host "    [2]: Microsoft Office 2021 Home and Business (Pro Plus Image)      " -ForegroundColor Cyan -NoNewLine
  Write-Host "║" -ForegroundColor Yellow
  Write-Host "              ║                                                                       ║" -ForegroundColor Yellow
  Write-Host "              ║" -ForegroundColor Yellow -NoNewLine
  Write-Host "    [3]: Microsoft Office 2021 Professional Plus                       " -ForegroundColor Cyan -NoNewLine
  Write-Host "║" -ForegroundColor Yellow
  Write-Host "              ║                                                                       ║" -ForegroundColor Yellow
  Write-Host "              ╠════════════════════════════" -NoNewline -ForegroundColor Yellow
  Write-Host "  Office 2024  " -ForegroundColor White -NoNewline
  Write-Host "════════════════════════════╣              " -ForegroundColor Yellow
  Write-Host "              ║                                                                       ║" -ForegroundColor Yellow
  Write-Host "              ║" -ForegroundColor Yellow -NoNewLine
  Write-Host "    [4]: Microsoft Office 2024 Home and Student                        " -ForegroundColor Cyan -NoNewLine
  Write-Host "║" -ForegroundColor Yellow
  Write-Host "              ║                                                                       ║" -ForegroundColor Yellow
  Write-Host "              ║" -ForegroundColor Yellow -NoNewLine
  Write-Host "    [5]: Microsoft Office 2024 Home and Business                       " -ForegroundColor Cyan -NoNewLine
  Write-Host "║" -ForegroundColor Yellow
  Write-Host "              ║                                                                       ║" -ForegroundColor Yellow
  Write-Host "              ║" -ForegroundColor Yellow -NoNewLine
  Write-Host "    [6]: Microsoft Office 2024 Professional Plus                       " -ForegroundColor Cyan -NoNewLine
  Write-Host "║" -ForegroundColor Yellow
  Write-Host "              ║                                                                       ║" -ForegroundColor Yellow
  Write-Host "              ║" -ForegroundColor Yellow -NoNewLine
  Write-Host "    [7]: Microsoft Office 2024 Home and Student (Pro Plus Image)       " -ForegroundColor Cyan -NoNewLine
  Write-Host "║" -ForegroundColor Yellow
  Write-Host "              ║                                                                       ║" -ForegroundColor Yellow
  Write-Host "              ║" -ForegroundColor Yellow -NoNewLine
  Write-Host "    [8]: Microsoft Office 2024 Home and Business (Pro Plus Image)      " -ForegroundColor Cyan -NoNewLine
  Write-Host "║" -ForegroundColor Yellow
  Write-Host "              ║                                                                       ║" -ForegroundColor Yellow
  Write-Host "              ╠════════════════════════════" -NoNewline -ForegroundColor Yellow
  Write-Host "  Office 365  " -ForegroundColor White -NoNewline
  Write-Host "═════════════════════════════╣              " -ForegroundColor Yellow
  Write-Host "              ║                                                                       ║" -ForegroundColor Yellow
  Write-Host "              ║" -ForegroundColor Yellow -NoNewLine
  Write-Host "    [9]: Microsoft Office 365 Single                                   " -ForegroundColor Cyan -NoNewLine
  Write-Host "║" -ForegroundColor Yellow
  Write-Host "              ║                                                                       ║" -ForegroundColor Yellow
  Write-Host "              ║" -ForegroundColor Yellow -NoNewLine
  Write-Host "    [10]: Microsoft Office 365 Family                                  " -ForegroundColor Cyan -NoNewLine
  Write-Host "║" -ForegroundColor Yellow
  Write-Host "              ║                                                                       ║" -ForegroundColor Yellow
  Write-Host "              ╠═══════════════════════════════════════════════════════════════════════╣" -ForegroundColor Yellow
  Write-Host "              ║                                                                       ║" -ForegroundColor Yellow
  Write-Host "              ║" -ForegroundColor Yellow -NoNewLine
  Write-Host "    [0]: Beenden                                       [11]: Readme    " -ForegroundColor Magenta -NoNewLine
  Write-Host "║" -ForegroundColor Yellow
  Write-Host "              ║                                                                       ║" -ForegroundColor Yellow
  Write-Host "              ╚═══════════════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
  Write-Host

  $actions = "0"
  while ($actions -notin "0..11") {
  $actions = Read-Host -Prompt '                  Was möchten Sie installieren?'
  
      if ($actions -in 0..11) {
          if ($actions -eq 0) {
              exit
            }

        #H&S2021
          if ($actions -eq 1) {
              $host.ui.RawUI.WindowTitle = "Installing Office 2021 Home and Student"
              Clear-Host
              Invoke-Command -ScriptBlock $Display
              Write-Host "══╦═══════════════════════════════════════════════════════════════════════════════════════════════╦══" -ForegroundColor Yellow
              Write-Host "  ║                                                                                               ║" -ForegroundColor Yellow
              Invoke-Command -ScriptBlock $MSHS21
              $Apps = @("Word", "Excel", "PowerPoint")
              Shortcut -Apps $Apps
              Invoke-Command -ScriptBlock $End
              exit
          }

        #H&B2021
          if ($actions -eq 2) {
              $host.ui.RawUI.WindowTitle = "Installing Office 2021 Home and Business"
              Clear-Host
              Invoke-Command -ScriptBlock $Display
              Write-Host "══╦═══════════════════════════════════════════════════════════════════════════════════════════════╦══" -ForegroundColor Yellow
              Write-Host "  ║                                                                                               ║" -ForegroundColor Yellow
              Invoke-Command -ScriptBlock $MSHB21
              $Apps = @("Word", "Excel", "PowerPoint", "Outlook (classic)")
              Shortcut -Apps $Apps 
              Invoke-Command -ScriptBlock $End
              exit
          }
        
        #Pro2021
          if ($actions -eq 3) {
              $host.ui.RawUI.WindowTitle = "Installing Office 2021 Professional Plus"
              Clear-Host
              Invoke-Command -ScriptBlock $Display
              Write-Host "══╦═══════════════════════════════════════════════════════════════════════════════════════════════╦══" -ForegroundColor Yellow
              Write-Host "  ║                                                                                               ║" -ForegroundColor Yellow
              Invoke-Command -ScriptBlock $MSPRO21
              $Apps = @("Word", "Excel", "PowerPoint", "Outlook (classic)", "Access", "Publisher")
              Shortcut -Apps $Apps 
              Invoke-Command -ScriptBlock $End
              exit
          }

        #H&S2024
          if ($actions -eq 4) {
            $host.ui.RawUI.WindowTitle = "Installing Office 2024 Home and Student"
            Clear-Host
            Invoke-Command -ScriptBlock $Display
            Write-Host "══╦═══════════════════════════════════════════════════════════════════════════════════════════════╦══" -ForegroundColor Yellow
            Write-Host "  ║                                                                                               ║" -ForegroundColor Yellow
            Invoke-Command -ScriptBlock $MSHS24
            $Apps = @("Word", "Excel", "PowerPoint")
            Shortcut -Apps $Apps
            Invoke-Command -ScriptBlock $End
            exit
          }

        #H&B2024
          if ($actions -eq 5) {
              $host.ui.RawUI.WindowTitle = "Installing Office 2024 Home and Business"
              Clear-Host
              Invoke-Command -ScriptBlock $Display
              Write-Host "══╦═══════════════════════════════════════════════════════════════════════════════════════════════╦══" -ForegroundColor Yellow
              Write-Host "  ║                                                                                               ║" -ForegroundColor Yellow
              Invoke-Command -ScriptBlock $MSHB21
              $Apps = @("Word", "Excel", "PowerPoint", "Outlook (classic)")
              Shortcut -Apps $Apps
              Invoke-Command -ScriptBlock $End
              exit
          } 

        #Pro2024
          if ($actions -eq 6) {
            $host.ui.RawUI.WindowTitle = "Installing Office 2024 Professional Plus"
            Clear-Host
            Invoke-Command -ScriptBlock $Display
            Write-Host "══╦═══════════════════════════════════════════════════════════════════════════════════════════════╦══" -ForegroundColor Yellow
            Write-Host "  ║                                                                                               ║" -ForegroundColor Yellow
            Invoke-Command -ScriptBlock $MSPRO24
            $Apps = @("Word", "Excel", "PowerPoint", "Outlook (classic)", "Access", "Publisher")
            Shortcut -Apps $Apps
            Invoke-Command -ScriptBlock $End
            exit
          }

        #H&SPro2024
          if ($actions -eq 7) {
              $host.ui.RawUI.WindowTitle = "Installing Office 2024 Home and Student"
              Clear-Host
              Invoke-Command -ScriptBlock $Display
              Write-Host "══╦═══════════════════════════════════════════════════════════════════════════════════════════════╦══" -ForegroundColor Yellow
              Write-Host "  ║                                                                                               ║" -ForegroundColor Yellow
              Invoke-Command -ScriptBlock $MSHS24P
              $Apps = @("Word", "Excel", "PowerPoint")
              Shortcut -Apps $Apps
              Invoke-Command -ScriptBlock $End
              exit
          } 

        #H&BPro2024
          if ($actions -eq 8) {
              $host.ui.RawUI.WindowTitle = "Installing Office 2024 Home and Business"
              Clear-Host
              Invoke-Command -ScriptBlock $Display
              Write-Host "══╦═══════════════════════════════════════════════════════════════════════════════════════════════╦══" -ForegroundColor Yellow
              Write-Host "  ║                                                                                               ║" -ForegroundColor Yellow
              Invoke-Command -ScriptBlock $MSHS24P
              $Apps = @("Word", "Excel", "PowerPoint", "Outlook (classic)")
              Shortcut -Apps $Apps
              Invoke-Command -ScriptBlock $End
              exit
          } 

        #365Single
          if ($actions -eq 9) {
              $host.ui.RawUI.WindowTitle = "Installing Office 365 Single"
              Clear-Host
              Invoke-Command -ScriptBlock $Display
              Write-Host "══╦═══════════════════════════════════════════════════════════════════════════════════════════════╦══" -ForegroundColor Yellow
              Write-Host "  ║                                                                                               ║" -ForegroundColor Yellow+
              Invoke-Command -ScriptBlock $MS365S
              $Apps = @("Word", "Excel", "PowerPoint", "Outlook (classic)", "Access", "Publisher")
              Shortcut -Apps $Apps
              Invoke-Command -ScriptBlock $End
              exit
          }

        #365Family
          if ($actions -eq 10) {
              $host.ui.RawUI.WindowTitle = "Installing Office 365 Family"
              Clear-Host
              Invoke-Command -ScriptBlock $Display
              Write-Host "══╦═══════════════════════════════════════════════════════════════════════════════════════════════╦══" -ForegroundColor Yellow
              Write-Host "  ║                                                                                               ║" -ForegroundColor Yellow
              Invoke-Command -ScriptBlock $MS365F
              $Apps = @("Word", "Excel", "PowerPoint", "Outlook (classic)", "Access", "Publisher")
              Shortcut -Apps $Apps
              Invoke-Command -ScriptBlock $End
              exit
          }      


          if ($actions -eq 11) {
            Start-Process "https://github.com/SD-ITLab/Silent_Office_Installer"
            menu
            }
            exit
            menu
        }
        else {
            menu
        }
    }
}
menu
