<#
.SYNOPSIS
  Professional Windows Update Manager - Enhanced visual experience with ASCII art, color, and interactive menus.

.DESCRIPTION
  WindowsUpdate-Manager v3 is an enhanced, production-ready PowerShell script that provides
  a professional, menu-driven console interface for Windows Update management with stunning visuals. Features include:

  Core Features (from v2):
    - Automated dependency checking and installation (PSWindowsUpdate, NuGet, PowerShell 7.5+)
    - Enhanced self-elevation with visual indicators
    - Comprehensive error handling and logging
    - Pre-flight system checks (disk space, connectivity, services)
    - System restore point creation before updates
    - KB exclusion lists
    - Update categorization (Security, Critical, Optional, Drivers)
    - Compliance reporting and audit trails

  NEW in v3 - Visual Enhancements:
    - ASCII art title with multiple styles (Standard, Slant, Cyberpunk)
    - Color gradients and 24-bit RGB support (PS7+)
    - Sound effects for key actions
    - Responsive layout that adapts to console size
    - Enhanced box-drawing characters for professional borders
    - Advanced status displays with visual indicators

  Update Operations:
    - View system status, reboot requirements, and last results
    - Scan for updates (Windows Update or Microsoft Update)
    - Install all updates or select specific updates
    - Hide/unhide updates with persistent exclusion lists
    - Uninstall updates by KB number
    - View and export update history
    - Reset Windows Update components
    - Remote installation via Invoke-WUJob (SYSTEM scheduled task)

  Configuration:
    - Persistent JSON configuration with validation
    - Visual settings (animations, sounds, themes)
    - Profile import/export
    - First-run wizard
    - Customizable logging options

.NOTES
  File Name      : WindowsUpdate-Manager_v3.ps1
  Author         : Ghostwheel
  Version        : 3.0.0
  Created        : 2026-01-19
  Last Modified  : 2026-01-19

  Requirements:
    - Windows PowerShell 5.1+ or PowerShell 7.x (7.5+ recommended for best visuals)
    - Administrator privileges (script will auto-elevate)
    - PSWindowsUpdate module (will auto-install if missing)
    - NuGet provider (will auto-install if missing)
    - Terminal with ANSI escape sequence support (Windows Terminal, PS7+)

  Links:
    - PSWindowsUpdate: https://www.powershellgallery.com/packages/PSWindowsUpdate
    - GitHub: https://github.com/mgajda83/PSWindowsUpdate

  License:
    This script is provided "as-is" without warranty. Use at your own risk.

.PARAMETER NoElevation
  Prevents automatic elevation to Administrator privileges. Use only if running pre-elevated
  or in constrained environments.

.PARAMETER ConfigPath
  Path to store configuration settings (JSON format).
  Default: $env:APPDATA\WindowsUpdateManager\config.json

.PARAMETER LogPath
  Path for transcript logging. If blank, logging is disabled by default but can be enabled
  via the settings menu.

.PARAMETER SkipDependencyCheck
  Skips automatic dependency checking and installation. Use only if dependencies are
  already confirmed to be installed.

.PARAMETER Silent
  Minimizes prompts for fully automated operation. Use with caution.

.EXAMPLE
  .\WindowsUpdate-Manager_v3.ps1
  Runs the script with default settings and auto-elevation.

.EXAMPLE
  .\WindowsUpdate-Manager_v3.ps1 -NoElevation -ConfigPath "C:\Config\wu.json"
  Runs without elevation using a custom configuration file path.

.EXAMPLE
  .\WindowsUpdate-Manager_v3.ps1 -LogPath "C:\Logs\WU.log"
  Runs with logging enabled to a specific path.
#>

[CmdletBinding()]
param(
  [Parameter(HelpMessage="Prevent automatic elevation")]
  [switch]$NoElevation,

  [Parameter(HelpMessage="Path to configuration file")]
  [ValidateNotNullOrEmpty()]
  [string]$ConfigPath = (Join-Path $env:APPDATA "WindowsUpdateManager\config.json"),

  [Parameter(HelpMessage="Path for transcript logging")]
  [string]$LogPath = "",

  [Parameter(HelpMessage="Skip dependency checks")]
  [switch]$SkipDependencyCheck,

  [Parameter(HelpMessage="Minimize user prompts")]
  [switch]$Silent
)

#region Script Variables
$script:Version = "3.0.0"
$script:TranscriptOn = $false
$script:LastScan = @()
$script:LastScanRaw = @()
$script:Config = $null
$script:IsElevated = $false
$script:ErrorLogPath = ""
$script:EscapeChar = [char]27
$script:SupportsANSI = $false
$script:ConsoleDimensions = @{Width=80; Height=24}
#endregion

#region ASCII Art & Visuals

# ASCII Art Titles
$script:AsciiArtTitles = @{
  "Standard" = @"
 _    _ _           _                      _   _           _       _
| |  | (_)         | |                    | | | |         | |     | |
| |  | |_ _ __   __| | _____      _____   | | | |_ __   __| | __ _| |_ ___
| |/\| | | '_ \ / _` |/ _ \ \ /\ / / __|  | | | | '_ \ / _` |/ _` | __/ _ \
\  /\  / | | | | (_| | (_) \ V  V /\__ \  | |_| | |_) | (_| | (_| | ||  __/
 \/  \/|_|_| |_|\__,_|\___/ \_/\_/ |___/   \___/| .__/ \__,_|\__,_|\__\___|
                                                 | |
                                                 |_|

╔╦╗┌─┐┌┐┌┌─┐┌─┐┌─┐┬─┐
║║║├─┤│││├─┤│ ┬├┤ ├┬┘
╩ ╩┴ ┴┘└┘┴ ┴└─┘└─┘┴└─
"@

  "Slant" = @"
 _       ___           __                     __  __          __      __
| |     / (_)___  ____/ /___ _      _____   / / / /___  ____/ /___ _/ /____
| | /| / / / __ \/ __  / __ \ | /| / / __|  / / / / __ \/ __  / __ `/ __/ _ \
| |/ |/ / / / / / /_/ / /_/ / |/ |/ /\__ \  / /_/ / /_/ / /_/ / /_/ / /_/  __/
|__/|__/_/_/ /_/\__,_/\____/|__/|__//___/  \____/ .___/\__,_/\__,_/\__/\___/
                                                /_/

╔╦╗┌─┐┌┐┌┌─┐┌─┐┌─┐┬─┐
║║║├─┤│││├─┤│ ┬├┤ ├┬┘
╩ ╩┴ ┴┘└┘┴ ┴└─┘└─┘┴└─
"@

  "Cyberpunk" = @"
╦ ╦┬┌┐┌┌┬┐┌─┐┬ ┬┌─┐  ╦ ╦┌─┐┌┬┐┌─┐┌┬┐┌─┐
║║║││││ │││ ││││└─┐  ║ ║├─┘ ││├─┤ │ ├┤
╚╩╝┴┘└┘─┴┘└─┘└┴┘└─┘  ╚═╝┴  ─┴┘┴ ┴ ┴ └─┘
╔╦╗┌─┐┌┐┌┌─┐┌─┐┌─┐┬─┐
║║║├─┤│││├─┤│ ┬├┤ ├┬┘
╩ ╩┴ ┴┘└┘┴ ┴└─┘└─┘┴└─
"@
}

function Get-ConsoleSize {
  try {
    $width = $Host.UI.RawUI.WindowSize.Width
    $height = $Host.UI.RawUI.WindowSize.Height

    if ($width -lt 1) { $width = 80 }
    if ($height -lt 1) { $height = 24 }

    return @{Width = $width; Height = $height}
  } catch {
    return @{Width = 80; Height = 24}
  }
}

function Set-OptimalWindowSize {
  param(
    [int]$RequiredWidth = 0,
    [int]$RequiredHeight = 0
  )
  <#
  .SYNOPSIS
    Auto-sizes the console window to fit the menu content optimally.
  #>
  try {
    # Skip if running in ISE or unsupported host
    if ($Host.Name -match "ISE" -or -not $Host.UI.RawUI) {
      $script:ConsoleDimensions = Get-ConsoleSize
      return
    }

    $raw = $Host.UI.RawUI

    # Optimal dimensions for the menu (width for descriptors, height for all options)
    $optimalWidth = if ($RequiredWidth -gt 0) { $RequiredWidth } else { 110 }
    $optimalHeight = if ($RequiredHeight -gt 0) { $RequiredHeight } else { 40 }
    $optimalWidth = [Math]::Max($optimalWidth, 80)
    $optimalHeight = [Math]::Max($optimalHeight, 24)

    $currentSize = $raw.WindowSize
    $maxSize = $raw.MaxPhysicalWindowSize
    $maxWidth = if ($maxSize.Width -gt 0) { $maxSize.Width } else { $currentSize.Width }
    $maxHeight = if ($maxSize.Height -gt 0) { $maxSize.Height } else { $currentSize.Height }

    # Calculate desired size (capped by max)
    $newWidth = [Math]::Min($optimalWidth, $maxWidth)
    $newHeight = [Math]::Min($optimalHeight, $maxHeight)

    # Expand buffer size first (must be >= window size)
    $bufferSize = $raw.BufferSize
    $targetBufferWidth = [Math]::Max($bufferSize.Width, $newWidth)
    $targetBufferHeight = [Math]::Max($bufferSize.Height, $newHeight)
    $bufferSize.Width = $targetBufferWidth
    $bufferSize.Height = $targetBufferHeight
    try { $raw.BufferSize = $bufferSize } catch { }

    # Clamp to actual buffer size if the resize was rejected
    $bufferSize = $raw.BufferSize
    $newWidth = [Math]::Min($newWidth, $bufferSize.Width)
    $newHeight = [Math]::Min($newHeight, $bufferSize.Height)

    # Only resize if needed and possible
    if ($currentSize.Width -ne $newWidth -or $currentSize.Height -ne $newHeight) {
      $windowSize = $raw.WindowSize
      $windowSize.Width = $newWidth
      $windowSize.Height = $newHeight
      try {
        $raw.WindowSize = $windowSize
      } catch {
        try { [Console]::SetWindowSize($newWidth, $newHeight) } catch { }
      }
    }

    # Expand buffer height for scrollback if supported
    try {
      $bufferSize = $raw.BufferSize
      if ($bufferSize.Height -lt 3000) {
        $bufferSize.Height = 3000
        $raw.BufferSize = $bufferSize
      }
    } catch { }

    # Center the window on screen if possible
    try {
      $windowPos = $raw.WindowPosition
      $windowPos.X = [Math]::Max(0, ($maxWidth - $newWidth) / 2)
      $windowPos.Y = [Math]::Max(0, ($maxHeight - $newHeight) / 2)
      $raw.WindowPosition = $windowPos
    } catch {
      # Centering might fail in some hosts, ignore
    }

    # Update script dimensions
    $script:ConsoleDimensions = Get-ConsoleSize
  } catch {
    $script:ConsoleDimensions = Get-ConsoleSize
  }
}

function Write-GradientText {
  param(
    [Parameter(Mandatory)]
    [string]$Text,

    [int]$StartR = 0, [int]$StartG = 255, [int]$StartB = 255,
    [int]$EndR = 0, [int]$EndG = 100, [int]$EndB = 255
  )

  # Check if gradients are enabled and supported
  if (-not $script:Config.Visual.EnableGradients -or -not $script:SupportsANSI) {
    Write-Host $Text -ForegroundColor Cyan
    return
  }

  # PowerShell 7+ supports 24-bit RGB
  if ($PSVersionTable.PSVersion.Major -ge 7) {
    $length = $Text.Length
    if ($length -le 1) {
      Write-Host $Text
      return
    }

    for ($i = 0; $i -lt $length; $i++) {
      $ratio = $i / ($length - 1)
      $r = [int]($StartR + ($EndR - $StartR) * $ratio)
      $g = [int]($StartG + ($EndG - $StartG) * $ratio)
      $b = [int]($StartB + ($EndB - $StartB) * $ratio)

      Write-Host "$script:EscapeChar[38;2;$r;$g;${b}m$($Text[$i])" -NoNewline
    }
    Write-Host "$script:EscapeChar[0m"
  } else {
    # Fallback for PowerShell 5.1
    Write-Host $Text -ForegroundColor Cyan
  }
}

function Write-RainbowText {
  param(
    [Parameter(Mandatory)]
    [string]$Text
  )

  if (-not $script:Config.Visual.EnableGradients -or -not $script:SupportsANSI) {
    Write-Host $Text -ForegroundColor Cyan
    return
  }

  if ($PSVersionTable.PSVersion.Major -ge 7) {
    $colors = @(
      @{R=255; G=0;   B=0},    # Red
      @{R=255; G=127; B=0},    # Orange
      @{R=255; G=255; B=0},    # Yellow
      @{R=0;   G=255; B=0},    # Green
      @{R=0;   G=0;   B=255},  # Blue
      @{R=75;  G=0;   B=130},  # Indigo
      @{R=148; G=0;   B=211}   # Violet
    )

    $length = $Text.Length
    for ($i = 0; $i -lt $length; $i++) {
      $colorIndex = [int](($i / $length) * ($colors.Count - 1))
      $c = $colors[$colorIndex]
      Write-Host "$script:EscapeChar[38;2;$($c.R);$($c.G);$($c.B)m$($Text[$i])" -NoNewline
    }
    Write-Host "$script:EscapeChar[0m"
  } else {
    Write-Host $Text -ForegroundColor Cyan
  }
}

function Play-Sound {
  param(
    [ValidateSet("Startup","Success","Error","Warning","Complete","Click")]
    [string]$SoundType
  )

  if (-not $script:Config.Sounds.$SoundType) {
    return
  }

  try {
    switch ($SoundType) {
      "Startup"  {
        [console]::beep(523,100)
        [console]::beep(659,100)
        [console]::beep(784,150)
      }
      "Success"  {
        [console]::beep(784,100)
        [console]::beep(1047,200)
      }
      "Error"    {
        [console]::beep(400,200)
        [console]::beep(300,300)
      }
      "Warning"  {
        [console]::beep(600,200)
      }
      "Complete" {
        [console]::beep(523,100)
        [console]::beep(659,100)
        [console]::beep(784,100)
        [console]::beep(1047,300)
      }
      "Click"    {
        [console]::beep(800,50)
      }
    }
  } catch {
    # Silently fail if beep not supported
  }
}

function Show-AsciiTitle {
  param([string]$Style = "Standard")

  $title = $script:AsciiArtTitles[$Style]
  if (-not $title) { $Style = "Standard"; $title = $script:AsciiArtTitles[$Style] }

  if ($script:Config.Visual.EnableGradients -and $PSVersionTable.PSVersion.Major -ge 7) {
    foreach ($line in ($title -split "`n")) {
      Write-GradientText -Text $line -StartR 0 -StartG 255 -StartB 255 -EndR 0 -EndG 150 -EndB 255
    }
  } else {
    Write-Host $title -ForegroundColor Cyan
  }
}

function New-BoxBorder {
  param(
    [Parameter(Mandatory)]
    [string]$Content,

    [string]$Title = "",

    [ValidateSet("Single","Double","Heavy","Rounded","ASCII")]
    [string]$Style = "Double",

    [int]$Width = 80,

    [ConsoleColor]$Color = 'Cyan'
  )

  # Box drawing character sets
  $borders = @{
    "Single"  = @{TL='┌'; TR='┐'; BL='└'; BR='┘'; H='─'; V='│'; TJ='┬'; BJ='┴'; LJ='├'; RJ='┤'; CJ='┼'}
    "Double"  = @{TL='╔'; TR='╗'; BL='╚'; BR='╝'; H='═'; V='║'; TJ='╦'; BJ='╩'; LJ='╠'; RJ='╣'; CJ='╬'}
    "Heavy"   = @{TL='┏'; TR='┓'; BL='┗'; BR='┛'; H='━'; V='┃'; TJ='┳'; BJ='┻'; LJ='┣'; RJ='┫'; CJ='╋'}
    "Rounded" = @{TL='╭'; TR='╮'; BL='╰'; BR='╯'; H='─'; V='│'; TJ='┬'; BJ='┴'; LJ='├'; RJ='┤'; CJ='┼'}
    "ASCII"   = @{TL='+'; TR='+'; BL='+'; BR='+'; H='-'; V='|'; TJ='+'; BJ='+'; LJ='+'; RJ='+'; CJ='+'}
  }

  $b = $borders[$Style]
  $innerWidth = $Width - 2

  # Top border
  $topBorder = $b.TL + ($b.H * $innerWidth) + $b.TR
  Write-Host $topBorder -ForegroundColor $Color

  # Title line if provided
  if ($Title) {
    $titleLen = $Title.Length
    $padding = $innerWidth - $titleLen - 2
    $leftPad = [Math]::Floor($padding / 2)
    $rightPad = $padding - $leftPad

    $titleLine = $b.V + (' ' * $leftPad) + $Title + (' ' * $rightPad) + $b.V
    Write-Host $titleLine -ForegroundColor $Color

    # Separator
    $separator = $b.LJ + ($b.H * $innerWidth) + $b.RJ
    Write-Host $separator -ForegroundColor $Color
  }

  # Content lines
  $contentLines = $Content -split "`n"
  foreach ($line in $contentLines) {
    $linelen = $line.Length
    $pad = $innerWidth - $linelen
    if ($pad -lt 0) {
      # Truncate if too long
      $line = $line.Substring(0, $innerWidth)
      $pad = 0
    }

    $contentLine = $b.V + $line + (' ' * $pad) + $b.V
    Write-Host $contentLine -ForegroundColor $Color
  }

  # Bottom border
  $bottomBorder = $b.BL + ($b.H * $innerWidth) + $b.BR
  Write-Host $bottomBorder -ForegroundColor $Color
}

function Show-ProgressBarEnhanced {
  param(
    [Parameter(Mandatory)]
    [string]$Activity,

    [Parameter(Mandatory)]
    [string]$Status,

    [int]$PercentComplete = 0,

    [int]$Total = 100,

    [int]$Current = 0,

    [string]$Speed = "",

    [string]$ETA = ""
  )

  # Use standard Write-Progress if animations disabled
  if (-not $script:Config.Visual.EnableAnimations) {
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
    return
  }

  $size = Get-ConsoleSize
  $barWidth = [Math]::Min(50, $size.Width - 40)

  $completed = [Math]::Floor($barWidth * $PercentComplete / 100)
  $remaining = $barWidth - $completed

  $bar = '█' * $completed + '░' * $remaining

  Write-Host ""
  Write-ColorOutput -Message "┌$('─' * ($barWidth + 30))┐" -ForegroundColor Cyan
  Write-Host "│ " -NoNewline -ForegroundColor Cyan
  Write-ColorOutput -Message $Activity -ForegroundColor White -NoNewline
  Write-Host (' ' * ($barWidth + 28 - $Activity.Length)) -NoNewline
  Write-Host "│" -ForegroundColor Cyan

  Write-Host "├$('─' * ($barWidth + 30))┤" -ForegroundColor Cyan

  Write-Host "│ " -NoNewline -ForegroundColor Cyan
  Write-ColorOutput -Message $Status -ForegroundColor Gray -NoNewline
  Write-Host (' ' * ($barWidth + 28 - $Status.Length)) -NoNewline
  Write-Host "│" -ForegroundColor Cyan

  Write-Host "│ " -NoNewline -ForegroundColor Cyan

  if ($PercentComplete -ge 75) {
    Write-Host $bar -NoNewline -ForegroundColor Green
  } elseif ($PercentComplete -ge 50) {
    Write-Host $bar -NoNewline -ForegroundColor Yellow
  } else {
    Write-Host $bar -NoNewline -ForegroundColor Cyan
  }

  Write-Host " $PercentComplete%" -NoNewline
  Write-Host (' ' * ($barWidth + 24 - $bar.Length - $PercentComplete.ToString().Length)) -NoNewline
  Write-Host "│" -ForegroundColor Cyan

  if ($Speed -or $ETA) {
    Write-Host "│ " -NoNewline -ForegroundColor Cyan
    $info = ""
    if ($Speed) { $info += "⚡ $Speed  " }
    if ($ETA) { $info += "⏱ ETA: $ETA" }
    Write-ColorOutput -Message $info -ForegroundColor Gray -NoNewline
    Write-Host (' ' * ($barWidth + 28 - $info.Length)) -NoNewline
    Write-Host "│" -ForegroundColor Cyan
  }

  Write-Host "└$('─' * ($barWidth + 30))┘" -ForegroundColor Cyan
}

#endregion

#region Console & UI Helpers (from v2, enhanced)

function Initialize-Colors {
  # Enable VT100 escape sequences for color support (Windows 10+)
  try {
    # Check ANSI support
    $script:SupportsANSI = $false

    if ($PSVersionTable.PSVersion.Major -ge 7) {
      $script:SupportsANSI = $true
    } elseif ($PSVersionTable.BuildVersion.Build -ge 10586) {
      $script:SupportsANSI = $true
    }

    if ($script:SupportsANSI) {
      $null = [System.Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    }
  } catch {
    $script:SupportsANSI = $false
  }

  # Update console dimensions
  $script:ConsoleDimensions = Get-ConsoleSize
}

function Write-Header {
  param(
    [Parameter(Mandatory)]
    [string]$Title,

    [Parameter(Mandatory)]
    [object]$Config
  )

  Clear-Host

  $line = "=" * $script:ConsoleDimensions.Width
  Write-ColorOutput -Message $line -ForegroundColor Cyan
  Write-ColorOutput -Message "  $Title" -ForegroundColor Cyan -Bold
  Write-ColorOutput -Message "  Version: $($script:Version)" -ForegroundColor Gray
  Write-ColorOutput -Message $line -ForegroundColor Cyan

  $computerDisplay = if ($Config.TargetComputers -and $Config.TargetComputers.Count -gt 0) {
    $Config.TargetComputers -join ", "
  } else {
    "LOCAL"
  }

  $sourceDisplay = if ($Config.UseMicrosoftUpdate) { "Microsoft Update" } else { "Windows Update" }

  Write-Host ("  Computer(s): {0}" -f $computerDisplay)
  Write-Host ("  Source     : {0}" -f $sourceDisplay)
  Write-Host ("  AutoReboot : {0}    IgnoreReboot: {1}    Verbose: {2}" -f $Config.AutoReboot, $Config.IgnoreReboot, $Config.Verbose)
  Write-Host ("  Logging    : {0}" -f $(if ($script:TranscriptOn) { "ON" } else { "OFF" }))

  if ($script:IsElevated) {
    Write-ColorOutput -Message "  Status     : Running as Administrator ✓" -ForegroundColor Green
  } else {
    Write-ColorOutput -Message "  Status     : Not elevated (some operations may fail)" -ForegroundColor Yellow
  }

  Write-ColorOutput -Message $line -ForegroundColor Cyan
}

function Write-ColorOutput {
  param(
    [Parameter(Mandatory)]
    [string]$Message,

    [Parameter()]
    [ConsoleColor]$ForegroundColor = [ConsoleColor]::White,

    [Parameter()]
    [ConsoleColor]$BackgroundColor,

    [Parameter()]
    [switch]$Bold,

    [Parameter()]
    [switch]$NoNewline
  )

  $params = @{
    Object = $Message
    ForegroundColor = $ForegroundColor
  }

  if ($BackgroundColor) {
    $params.BackgroundColor = $BackgroundColor
  }

  if ($NoNewline) {
    $params.NoNewline = $true
  }

  Write-Host @params
}

function Write-Info {
  param([Parameter(Mandatory)][string]$Message)
  Write-ColorOutput -Message ("[ℹ] {0}" -f $Message) -ForegroundColor Cyan
}

function Write-Ok {
  param([Parameter(Mandatory)][string]$Message)
  Write-ColorOutput -Message ("[✓] {0}" -f $Message) -ForegroundColor Green
}

function Write-Warn {
  param([Parameter(Mandatory)][string]$Message)
  Write-ColorOutput -Message ("[⚠] {0}" -f $Message) -ForegroundColor Yellow
}

function Write-Err {
  param([Parameter(Mandatory)][string]$Message)
  Write-ColorOutput -Message ("[✗] {0}" -f $Message) -ForegroundColor Red
  Write-ErrorLog -Message $Message

  if ($script:Config.Sounds.Error) {
    Play-Sound -SoundType "Error"
  }
}

function Write-Success {
  param([Parameter(Mandatory)][string]$Message)
  Write-ColorOutput -Message ("[✓] {0}" -f $Message) -ForegroundColor Green

  if ($script:Config.Sounds.Success) {
    Play-Sound -SoundType "Success"
  }
}

function Write-Progress-Enhanced {
  param(
    [Parameter(Mandatory)]
    [string]$Activity,

    [Parameter(Mandatory)]
    [string]$Status,

    [Parameter()]
    [int]$PercentComplete = -1
  )

  if ($PercentComplete -ge 0) {
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
  } else {
    Write-Progress -Activity $Activity -Status $Status
  }
}

function Pause-UI {
  param([string]$Message = "Press Enter to continue...")
  Write-Host ""
  [void](Read-Host $Message)
}

function Read-YesNo {
  param(
    [Parameter(Mandatory)]
    [string]$Prompt,

    [switch]$DefaultYes
  )

  if ($Silent) { return [bool]$DefaultYes }

  $suffix = if ($DefaultYes) { " [Y/n]" } else { " [y/N]" }
  $answer = Read-Host ($Prompt + $suffix)

  if ([string]::IsNullOrWhiteSpace($answer)) {
    return [bool]$DefaultYes
  }

  return ($answer.Trim().ToLowerInvariant() -in @("y","yes"))
}

function Show-Banner {
  Clear-Host

  # Play startup sound
  Play-Sound -SoundType "Startup"

  # Show ASCII art title
  $style = $script:Config.Visual.AsciiArtStyle
  if (-not $style) { $style = "Cyberpunk" }

  Show-AsciiTitle -Style $style

  Write-Host ""
  Write-ColorOutput -Message "  Professional Update Management Tool" -ForegroundColor Gray
  Write-ColorOutput -Message "  Version $($script:Version)" -ForegroundColor Gray
  Write-Host ""
}

#endregion

#region Error Logging (unchanged from v2)

function Initialize-ErrorLog {
  try {
    $logDir = Join-Path $env:APPDATA "WindowsUpdateManager\logs"
    if (-not (Test-Path $logDir)) {
      New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    $script:ErrorLogPath = Join-Path $logDir ("ErrorLog_{0}.txt" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

    Add-Content -Path $script:ErrorLogPath -Value ("=" * 80) -ErrorAction SilentlyContinue
    Add-Content -Path $script:ErrorLogPath -Value "Windows Update Manager v$($script:Version) - Error Log" -ErrorAction SilentlyContinue
    Add-Content -Path $script:ErrorLogPath -Value ("Started: {0}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) -ErrorAction SilentlyContinue
    Add-Content -Path $script:ErrorLogPath -Value ("=" * 80) -ErrorAction SilentlyContinue
    Add-Content -Path $script:ErrorLogPath -Value "" -ErrorAction SilentlyContinue
  } catch {
    # Silent fail on log initialization
  }
}

function Write-ErrorLog {
  param(
    [Parameter(Mandatory)]
    [string]$Message,

    [Parameter()]
    [System.Management.Automation.ErrorRecord]$ErrorRecord
  )

  try {
    if ([string]::IsNullOrEmpty($script:ErrorLogPath)) { return }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[{0}] {1}" -f $timestamp, $Message

    Add-Content -Path $script:ErrorLogPath -Value $logEntry -ErrorAction SilentlyContinue

    if ($ErrorRecord) {
      Add-Content -Path $script:ErrorLogPath -Value ("  Exception: {0}" -f $ErrorRecord.Exception.Message) -ErrorAction SilentlyContinue
      Add-Content -Path $script:ErrorLogPath -Value ("  Category : {0}" -f $ErrorRecord.CategoryInfo.Category) -ErrorAction SilentlyContinue
      Add-Content -Path $script:ErrorLogPath -Value "" -ErrorAction SilentlyContinue
    }
  } catch {
    # Silent fail on error logging
  }
}

#endregion

#region Admin & Environment (unchanged from v2)

function Test-IsAdmin {
  try {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  } catch {
    Write-ErrorLog -Message "Failed to check admin status" -ErrorRecord $_
    return $false
  }
}

function Ensure-Elevation {
  param([switch]$Skip)

  if ($Skip) {
    Write-Warn "Skipping elevation check. Some operations may fail without admin rights."
    return
  }

  if (Test-IsAdmin) {
    $script:IsElevated = $true
    return
  }

  Write-Warn "Administrator privileges required. Attempting to elevate..."

  try {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = if ($PSVersionTable.PSVersion.Major -ge 7) { "pwsh.exe" } else { "powershell.exe" }

    $args = @(
      "-NoProfile",
      "-ExecutionPolicy", "Bypass",
      "-File", ('"{0}"' -f $PSCommandPath)
    )

    if ($NoElevation) { $args += "-NoElevation" }
    if ($ConfigPath) { $args += @("-ConfigPath", ('"{0}"' -f $ConfigPath)) }
    if ($LogPath) { $args += @("-LogPath", ('"{0}"' -f $LogPath)) }
    if ($SkipDependencyCheck) { $args += "-SkipDependencyCheck" }
    if ($Silent) { $args += "-Silent" }

    $psi.Arguments = ($args -join " ")
    $psi.Verb = "runas"
    $psi.UseShellExecute = $true

    $process = [Diagnostics.Process]::Start($psi)

    if ($null -ne $process) {
      Write-Success "Elevated instance started. This window will close."
      Start-Sleep -Seconds 2
      exit 0
    }
  } catch {
    Write-Err "Elevation failed or was cancelled."
    Write-ErrorLog -Message "Elevation failed" -ErrorRecord $_
    Write-Warn "Some operations require administrator rights and may not work."
    Start-Sleep -Seconds 3
  }
}

function Test-InternetConnectivity {
  try {
    $result = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue
    return $result
  } catch {
    return $false
  }
}

function Get-FreeDiskSpace {
  param([string]$Drive = "C:")

  try {
    $disk = Get-PSDrive -Name $Drive.TrimEnd(':') -ErrorAction Stop
    return [math]::Round($disk.Free / 1GB, 2)
  } catch {
    return 0
  }
}

function Test-WindowsUpdateService {
  try {
    $service = Get-Service -Name "wuauserv" -ErrorAction Stop
    return ($service.Status -eq 'Running')
  } catch {
    return $false
  }
}

#endregion

#region Pre-flight Checks (unchanged from v2)

function Invoke-PreFlightChecks {
  param([switch]$Detailed)

  Write-Info "Running pre-flight system checks..."
  $allPassed = $true

  # Check 1: Internet connectivity
  Write-Host "  Checking internet connectivity..." -NoNewline
  if (Test-InternetConnectivity) {
    Write-ColorOutput -Message " ✓" -ForegroundColor Green
  } else {
    Write-ColorOutput -Message " ✗" -ForegroundColor Red
    Write-Warn "No internet connectivity detected. Update operations will fail."
    $allPassed = $false
  }

  # Check 2: Disk space
  Write-Host "  Checking disk space..." -NoNewline
  $freeSpace = Get-FreeDiskSpace -Drive "C:"
  if ($freeSpace -gt 10) {
    Write-ColorOutput -Message (" ✓ ({0:N2} GB free)" -f $freeSpace) -ForegroundColor Green
  } elseif ($freeSpace -gt 5) {
    Write-ColorOutput -Message (" ⚠ ({0:N2} GB free - Low)" -f $freeSpace) -ForegroundColor Yellow
  } else {
    Write-ColorOutput -Message (" ✗ ({0:N2} GB free - Critical)" -f $freeSpace) -ForegroundColor Red
    Write-Warn "Insufficient disk space. At least 10GB recommended for updates."
    $allPassed = $false
  }

  # Check 3: Windows Update service
  Write-Host "  Checking Windows Update service..." -NoNewline
  if (Test-WindowsUpdateService) {
    Write-ColorOutput -Message " ✓" -ForegroundColor Green
  } else {
    Write-ColorOutput -Message " ⚠" -ForegroundColor Yellow
    Write-Warn "Windows Update service is not running. Attempting to start..."
    try {
      Start-Service -Name "wuauserv" -ErrorAction Stop
      Write-Success "Windows Update service started successfully."
    } catch {
      Write-Err "Failed to start Windows Update service."
      $allPassed = $false
    }
  }

  # Check 4: PowerShell version
  if ($Detailed) {
    Write-Host "  PowerShell version..." -NoNewline
    $psVersion = $PSVersionTable.PSVersion
    if ($psVersion.Major -ge 7 -and $psVersion.Minor -ge 5) {
      Write-ColorOutput -Message (" ✓ ({0})" -f $psVersion) -ForegroundColor Green
    } elseif ($psVersion.Major -ge 5) {
      Write-ColorOutput -Message (" ⚠ ({0} - Consider upgrading)" -f $psVersion) -ForegroundColor Yellow
    } else {
      Write-ColorOutput -Message (" ✗ ({0} - Unsupported)" -f $psVersion) -ForegroundColor Red
      $allPassed = $false
    }
  }

  Write-Host ""

  if ($allPassed) {
    Write-Success "All pre-flight checks passed."
  } else {
    Write-Warn "Some pre-flight checks failed. Proceed with caution."
  }

  return $allPassed
}

#endregion

#region Dependency Management (unchanged from v2)

function Test-NuGetProvider {
  try {
    $nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
    return ($null -ne $nuget)
  } catch {
    return $false
  }
}

function Install-NuGetProvider {
  Write-Info "Installing NuGet package provider..."

  try {
    # Ensure TLS 1.2 for PowerShell Gallery
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop | Out-Null
    Write-Success "NuGet provider installed successfully."
    return $true
  } catch {
    Write-Err "Failed to install NuGet provider: $($_.Exception.Message)"
    Write-ErrorLog -Message "NuGet installation failed" -ErrorRecord $_
    return $false
  }
}

function Get-InstalledPowerShellVersion {
  return $PSVersionTable.PSVersion
}

function Test-PowerShell7Available {
  try {
    $ps7Path = if ($env:ProgramFiles) {
      Join-Path $env:ProgramFiles "PowerShell\7\pwsh.exe"
    } else {
      $null
    }

    if ($ps7Path -and (Test-Path $ps7Path)) {
      return $true
    }

    $pwsh = Get-Command pwsh.exe -ErrorAction SilentlyContinue
    return ($null -ne $pwsh)
  } catch {
    return $false
  }
}

function Offer-PowerShell7Installation {
  $currentVersion = Get-InstalledPowerShellVersion

  # Only offer if running Windows PowerShell 5.1
  if ($currentVersion.Major -ge 7) {
    if ($currentVersion.Minor -lt 5) {
      Write-Info "You're running PowerShell 7.$($currentVersion.Minor). Consider upgrading to 7.5+ for best performance."
    }
    return
  }

  # Check if PS7 is already installed but not running
  if (Test-PowerShell7Available) {
    Write-Info "PowerShell 7 is installed. You can run this script with 'pwsh.exe' for better performance and visuals."
    return
  }

  Write-Host ""
  Write-ColorOutput -Message "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
  Write-ColorOutput -Message "  PowerShell 7.5+ Recommended" -ForegroundColor Cyan
  Write-ColorOutput -Message "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
  Write-Host ""
  Write-Host "You're currently running Windows PowerShell $currentVersion"
  Write-Host ""
  Write-Host "Benefits of upgrading to PowerShell 7.5+ for v3:"
  Write-ColorOutput -Message "  ✓ Better performance and faster execution" -ForegroundColor Green
  Write-ColorOutput -Message "  ✓ 24-bit RGB color support (gradients!)" -ForegroundColor Green
  Write-ColorOutput -Message "  ✓ Enhanced ANSI escape sequence support" -ForegroundColor Green
  Write-ColorOutput -Message "  ✓ Improved error handling and debugging" -ForegroundColor Green
  Write-ColorOutput -Message "  ✓ Enhanced JSON and data processing" -ForegroundColor Green
  Write-ColorOutput -Message "  ✓ Parallel processing capabilities" -ForegroundColor Green
  Write-ColorOutput -Message "  ✓ Long-term support from Microsoft" -ForegroundColor Green
  Write-Host ""
  Write-Host "Note: Windows PowerShell 5.1 will still work, but PS 7.5+ provides the best visual experience."
  Write-Host ""

  if (Read-YesNo "Would you like to install PowerShell 7.5+?" -DefaultYes) {
    Install-PowerShell7
  } else {
    Write-Info "Continuing with Windows PowerShell $currentVersion (some visual features may be limited)"
  }
}

function Install-PowerShell7 {
  Write-Info "Checking for winget availability..."

  # Try winget first (recommended method)
  $winget = Get-Command winget -ErrorAction SilentlyContinue

  if ($winget) {
    Write-Info "Installing PowerShell 7 using winget..."
    try {
      $result = & winget install --id Microsoft.PowerShell --source winget --accept-package-agreements --accept-source-agreements 2>&1

      if ($LASTEXITCODE -eq 0) {
        Write-Success "PowerShell 7 installed successfully via winget!"
        Write-Info "Please restart this script using 'pwsh.exe' to use PowerShell 7."
        Write-Info "Command: pwsh.exe -File `"$PSCommandPath`""
        Pause-UI "Press Enter to continue with current session..."
        return $true
      } else {
        Write-Warn "Winget installation completed with warnings. PS7 may be installed."
      }
    } catch {
      Write-Warn "Winget installation encountered an error. Trying alternative method..."
    }
  }

  # Fallback: Direct MSI download
  Write-Info "Installing PowerShell 7 using MSI installer..."
  Write-Info "Opening download page in browser. Please download and install manually."

  try {
    $downloadUrl = "https://aka.ms/install-powershell"
    Start-Process $downloadUrl
    Write-Info "Download page opened. After installing, restart this script with 'pwsh.exe'."
    Pause-UI
  } catch {
    Write-Err "Failed to open download page. Please visit: https://aka.ms/install-powershell"
    Write-ErrorLog -Message "Failed to open PS7 download page" -ErrorRecord $_
  }

  return $false
}

function Ensure-PSWindowsUpdate {
  if (Get-Module -ListAvailable -Name PSWindowsUpdate) {
    # Check version
    $module = Get-Module -ListAvailable -Name PSWindowsUpdate | Sort-Object Version -Descending | Select-Object -First 1
    Write-Info "PSWindowsUpdate version $($module.Version) is installed."
    return $true
  }

  Write-Warn "PSWindowsUpdate module is not installed."
  Write-Info "This module is required for Windows Update management."

  if (-not (Read-YesNo "Install PSWindowsUpdate from PowerShell Gallery now?" -DefaultYes)) {
    return $false
  }

  # Ensure prerequisites
  if (-not (Test-NuGetProvider)) {
    if (-not (Install-NuGetProvider)) {
      return $false
    }
  }

  # Ensure TLS 1.2
  try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  } catch {
    Write-Warn "Could not set TLS 1.2. Installation may fail on older systems."
  }

  Write-Info "Installing PSWindowsUpdate module..."
  Write-Info "This may take a few moments..."

  try {
    Install-Module -Name PSWindowsUpdate -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    Write-Success "PSWindowsUpdate module installed successfully."
    return $true
  } catch {
    Write-Err "Failed to install PSWindowsUpdate: $($_.Exception.Message)"
    Write-ErrorLog -Message "PSWindowsUpdate installation failed" -ErrorRecord $_
    Write-Info "Try running as administrator or install manually:"
    Write-Info "  Install-Module -Name PSWindowsUpdate -Scope CurrentUser"
    return $false
  }
}

function Import-PSWindowsUpdate {
  try {
    Import-Module PSWindowsUpdate -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
    Write-Success "PSWindowsUpdate module imported successfully."
    return $true
  } catch {
    Write-Err "Failed to import PSWindowsUpdate: $($_.Exception.Message)"
    Write-ErrorLog -Message "PSWindowsUpdate import failed" -ErrorRecord $_
    return $false
  }
}

function Ensure-MicrosoftUpdateService {
  param([switch]$Force)

  if (-not $script:Config.UseMicrosoftUpdate) { return }

  try {
    $services = Get-WUServiceManager -ErrorAction Stop
    $hasMU = $false

    foreach ($service in @($services)) {
      if ($service.Name -match "Microsoft Update") {
        $hasMU = $true
        break
      }
    }

    if ($hasMU -and -not $Force) {
      return
    }

    Write-Info "Registering Microsoft Update service manager..."
    Add-WUServiceManager -MicrosoftUpdate -ErrorAction Stop | Out-Null
    Write-Success "Microsoft Update service manager registered."
  } catch {
    Write-Warn "Could not verify/register Microsoft Update service: $($_.Exception.Message)"
    Write-ErrorLog -Message "Microsoft Update service registration failed" -ErrorRecord $_
    Write-Info "You can still scan using the default Windows Update service."
  }
}

function Initialize-Dependencies {
  if ($SkipDependencyCheck) {
    Write-Info "Dependency check skipped by parameter."
    return $true
  }

  Write-Info "Checking dependencies..."
  Write-Host ""

  # Check PowerShell version
  $psVersion = Get-InstalledPowerShellVersion
  Write-Host ("  PowerShell Version: {0}.{1}.{2}" -f $psVersion.Major, $psVersion.Minor, $psVersion.Build)

  # Check NuGet
  Write-Host "  NuGet Provider..." -NoNewline
  if (Test-NuGetProvider) {
    Write-ColorOutput -Message " ✓" -ForegroundColor Green
  } else {
    Write-ColorOutput -Message " ✗ (Not installed)" -ForegroundColor Yellow
    Write-Host ""
    Write-Info "NuGet provider is required to install PowerShell modules."
    if (Read-YesNo "Install NuGet provider? (Required for PSWindowsUpdate module)" -DefaultYes) {
      Write-Info "Installing NuGet provider..."
      if (-not (Install-NuGetProvider)) {
        Write-Err "Failed to install NuGet provider. Cannot continue."
        return $false
      }
      Write-Success "NuGet provider installed successfully."
    } else {
      Write-Err "NuGet provider is required. Cannot continue."
      return $false
    }
  }

  # Check PSWindowsUpdate
  Write-Host "  PSWindowsUpdate Module..." -NoNewline
  if (Get-Module -ListAvailable -Name PSWindowsUpdate) {
    Write-ColorOutput -Message " ✓" -ForegroundColor Green
  } else {
    Write-ColorOutput -Message " ✗ (Not installed)" -ForegroundColor Yellow
    Write-Host ""
    Write-Info "PSWindowsUpdate module provides Windows Update management cmdlets."
    if (Read-YesNo "Install PSWindowsUpdate module? (Required for this script)" -DefaultYes) {
      Write-Info "Installing PSWindowsUpdate module..."
      if (-not (Ensure-PSWindowsUpdate)) {
        Write-Err "Failed to install PSWindowsUpdate module. Cannot continue."
        return $false
      }
      Write-Success "PSWindowsUpdate module installed successfully."
    } else {
      Write-Err "PSWindowsUpdate module is required. Cannot continue."
      return $false
    }
  }

  Write-Host ""
  Write-Success "All dependencies satisfied."
  Write-Host ""

  # Offer PowerShell 7 upgrade
  if (-not $Silent) {
    Offer-PowerShell7Installation
  }

  # Import PSWindowsUpdate
  if (-not (Import-PSWindowsUpdate)) {
    return $false
  }

  return $true
}

#endregion

#region Configuration Management (enhanced for v3)

function Ensure-ConfigDirectory {
  $dir = Split-Path -Parent $ConfigPath
  if (-not (Test-Path $dir)) {
    try {
      New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
    } catch {
      Write-ErrorLog -Message "Failed to create config directory" -ErrorRecord $_
      throw
    }
  }
}

function Get-DefaultConfig {
  return [pscustomobject]@{
    UseMicrosoftUpdate = $true
    AutoReboot         = $false
    IgnoreReboot       = $false
    Verbose            = $false
    TargetComputers    = @()
    CreateRestorePoint = $true
    ExcludedKBs        = @()
    AutoAcceptEULA     = $false
    Visual = [pscustomobject]@{
      AsciiArtStyle      = "Cyberpunk"
      EnableAnimations   = $true
      EnableSounds       = $false
      AnimationSpeed     = "Normal"
      ColorTheme         = "Default"
      ShowTransitions    = $false
      BorderStyle        = "Double"
      ShowLiveDashboard  = $false
      EnableGradients    = $true
      PerformanceMode    = $false
    }
    Sounds = [pscustomobject]@{
      Startup  = $false
      Success  = $false
      Error    = $false
      Warning  = $false
      Complete = $false
      MenuClick = $false
    }
  }
}

function Load-Config {
  Ensure-ConfigDirectory

  if (Test-Path $ConfigPath) {
    try {
      $raw = Get-Content -Path $ConfigPath -Raw -ErrorAction Stop
      $cfg = $raw | ConvertFrom-Json -ErrorAction Stop
      return (Normalize-Config $cfg)
    } catch {
      Write-Warn "Config file exists but couldn't be read. Creating default config."
      Write-ErrorLog -Message "Failed to load config" -ErrorRecord $_
    }
  }

  $default = Get-DefaultConfig
  Save-Config $default
  return $default
}

function Normalize-Config {
  param($cfg)

  # Ensure all fields exist with defaults
  $default = Get-DefaultConfig

  foreach ($prop in $default.PSObject.Properties) {
    if ($null -eq $cfg.PSObject.Properties[$prop.Name]) {
      $cfg | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value -Force
    }
  }

  # Ensure nested Visual object exists
  if (-not $cfg.Visual) {
    $cfg | Add-Member -NotePropertyName "Visual" -NotePropertyValue $default.Visual -Force
  } else {
    foreach ($prop in $default.Visual.PSObject.Properties) {
      if ($null -eq $cfg.Visual.PSObject.Properties[$prop.Name]) {
        $cfg.Visual | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value -Force
      }
    }
  }

  # Ensure nested Sounds object exists
  if (-not $cfg.Sounds) {
    $cfg | Add-Member -NotePropertyName "Sounds" -NotePropertyValue $default.Sounds -Force
  } else {
    foreach ($prop in $default.Sounds.PSObject.Properties) {
      if ($null -eq $cfg.Sounds.PSObject.Properties[$prop.Name]) {
        $cfg.Sounds | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value -Force
      }
    }
  }

  # Ensure arrays
  if ($cfg.TargetComputers -isnot [System.Array]) {
    $cfg.TargetComputers = @($cfg.TargetComputers)
  }

  if ($cfg.ExcludedKBs -isnot [System.Array]) {
    $cfg.ExcludedKBs = @($cfg.ExcludedKBs)
  }

  return $cfg
}

function Save-Config {
  param($cfg)

  try {
    Ensure-ConfigDirectory
    $cfg | ConvertTo-Json -Depth 10 | Set-Content -Path $ConfigPath -Encoding UTF8 -Force -ErrorAction Stop
  } catch {
    Write-Err "Failed to save configuration: $($_.Exception.Message)"
    Write-ErrorLog -Message "Failed to save config" -ErrorRecord $_
  }
}

function Show-FirstRunWizard {
  Show-Banner

  Write-ColorOutput -Message "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
  Write-ColorOutput -Message "  First Run Configuration Wizard" -ForegroundColor Cyan
  Write-ColorOutput -Message "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
  Write-Host ""
  Write-Info "Let's configure your preferences for Windows Update management."
  Write-Host ""

  # Update Source
  Write-Host "Update Source:"
  Write-Host "  [1] Microsoft Update (includes Windows + Microsoft products)"
  Write-Host "      Include Office and other Microsoft apps"
  Write-Host "  [2] Windows Update only"
  Write-Host "      Only Windows OS updates"
  $sourceChoice = Read-Host "Choice [1]"
  $config.UseMicrosoftUpdate = ($sourceChoice -ne "2")

  Write-Host ""

  # Auto Reboot
  $config.AutoReboot = Read-YesNo "Automatically reboot after updates if required?" -DefaultYes:$false

  # Create Restore Point
  $config.CreateRestorePoint = Read-YesNo "Create system restore point before installing updates?" -DefaultYes:$true

  Write-Host ""

  # Visual preferences
  Write-Info "Visual Preferences:"
  $config.Visual.EnableSounds = Read-YesNo "Enable sound effects?" -DefaultYes:$true
  $config.Visual.EnableGradients = Read-YesNo "Enable color gradients (PS7+ only)?" -DefaultYes:$true

  Write-Host ""
  Write-Success "Configuration saved!"
  Write-Host ""

  Save-Config $config
  Pause-UI
}

#endregion
#endregion Configuration Management

#region Selection UI

function Test-OutGridViewAvailable {
    try {
        Get-Command Out-GridView -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Select-FromList {
    param(
        [Parameter(Mandatory)]
        [array]$Items,
        [Parameter(Mandatory)]
        [string]$Property,
        [string]$Title = "Select Items"
    )

    if ($Items.Count -eq 0) {
        Write-Warning "No items available for selection."
        return @()
    }

    if (Test-OutGridViewAvailable) {
        Write-Info "Opening selection window (Out-GridView)..."
        $selected = $Items | Out-GridView -Title $Title -OutputMode Multiple
        return $selected
    } else {
        Write-Warning "Out-GridView not available. Using console selection."
        Write-Host ""
        Write-Host "Available items:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $Items.Count; $i++) {
            Write-Host "  [$($i+1)] $($Items[$i].$Property)"
        }
        Write-Host ""
        $input = Read-Host "Enter numbers separated by commas (e.g., 1,3,5) or 'all'"

        if ($input -eq 'all') {
            return $Items
        }

        $indices = $input -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ - 1 }
        $selected = @()
        foreach ($idx in $indices) {
            if ($idx -ge 0 -and $idx -lt $Items.Count) {
                $selected += $Items[$idx]
            }
        }
        return $selected
    }
}

#endregion Selection UI

#region Update Object Mapping

function Get-UpdateDisplay {
    param([Parameter(Mandatory)]$Update)

    $severity = if ($Update.MsrcSeverity) { $Update.MsrcSeverity } else { "N/A" }
    $kb = "N/A"
    if ($Update.KBArticleIDs -and $Update.KBArticleIDs.Count -gt 0) {
        $kb = "KB" + $Update.KBArticleIDs[0]
    }

    return [PSCustomObject]@{
        Title       = $Update.Title
        KB          = $kb
        Severity    = $severity
        Size        = if ($Update.MaxDownloadSize) { "{0:N2} MB" -f ($Update.MaxDownloadSize / 1MB) } else { "N/A" }
        RebootReq   = $Update.RebootRequired
        UpdateObj   = $Update
    }
}

#endregion Update Object Mapping

#region Parameter Helpers

function Get-TargetsParam {
    if ($script:Config.TargetComputers -and $script:Config.TargetComputers.Count -gt 0) {
        return @{ ComputerName = $script:Config.TargetComputers }
    } else {
        return @{}
    }
}

function Get-VerboseParam {
    if ($script:Config.Verbose) {
        return @{ Verbose = $true }
    } else {
        return @{}
    }
}

#endregion Parameter Helpers

#region System Restore Point

function New-SystemRestorePoint {
    param([string]$Description = "Windows Update Manager")

    if (-not $script:Config.CreateRestorePoint) {
        Write-Info "Restore point creation disabled in settings."
        return
    }

    if (-not $script:IsElevated) {
        Write-Warning "Restore point creation requires Administrator privileges."
        return
    }

    try {
        Write-Info "Creating system restore point..."
        Checkpoint-Computer -Description $Description -RestorePointType MODIFY_SETTINGS -ErrorAction Stop
        Write-Success "Restore point created successfully."
        Play-Sound -SoundType Success
    } catch {
        Write-Err "Failed to create restore point: $_"
    }
}

#endregion System Restore Point

#region Action Functions


function Action-ShowStatus {
    Write-Header "Windows Update Status"

    try {
        $targets = Get-TargetsParam
        $verbose = Get-VerboseParam

        Write-Info "Retrieving Windows Update status..."
        $status = Get-WUInstallerStatus @targets @verbose

        Write-Host ""
        Write-Host "Installer Status:" -ForegroundColor Cyan
        $status | Format-List

        Write-Host ""
        Write-Host "Reboot Status:" -ForegroundColor Cyan
        $rebootStatus = Get-WURebootStatus @targets @verbose
        if ($rebootStatus) {
            Write-Warning "Reboot is required!"
        } else {
            Write-Success "No reboot required."
        }

        Write-Host ""
        Write-Host "Last Update Results:" -ForegroundColor Cyan
        $lastResults = Get-WULastResults @targets @verbose
        $lastResults | Format-List

        Write-Host ""
        Write-Host "Service Managers:" -ForegroundColor Cyan
        $serviceManagers = Get-WUServiceManager @targets @verbose
        $serviceManagers | Format-Table -AutoSize

    } catch {
        Write-Err "Failed to retrieve status: $_"
        Write-ErrorLog $_
    }

    Pause
}

function Prompt-SearchFilters {
  $filters = [ordered]@{}

  Write-Host ""
  Write-Info "Optional: Apply search filters (press Enter to skip)"
  Write-Host ""

  $title = Read-Host "Title regex filter"
  if (-not [string]::IsNullOrWhiteSpace($title)) {
    $filters.Title = $title
  }

  $kb = Read-Host "KB filter (e.g. KB5031234)"
  if (-not [string]::IsNullOrWhiteSpace($kb)) {
    $filters.KBArticleID = $kb.Trim()
  }

  # Category filter
  Write-Host ""
  Write-Host "Filter by category:"
  Write-Host "  [1] All updates"
  Write-Host "      No category filter"
  Write-Host "  [2] Security updates only"
  Write-Host "      Security Updates category"
  Write-Host "  [3] Critical updates only"
  Write-Host "      Critical Updates category"
  Write-Host "  [4] Drivers only"
  Write-Host "      Driver updates only"
  $catChoice = Read-Host "Choice [1]"

  switch ($catChoice) {
    "2" { $filters.Category = "Security Updates" }
    "3" { $filters.Category = "Critical Updates" }
    "4" { $filters.Category = "Drivers" }
  }

  return $filters
}

function Action-ScanUpdates {
  Write-Header "Windows Update Manager — Scan for Updates" $script:Config

  # Pre-flight checks
  if (-not (Invoke-PreFlightChecks)) {
    Write-Warning "Pre-flight checks failed. Some operations may not work."
    if (-not (Read-YesNo "Continue anyway?")) {
      return
    }
  }

  Ensure-MicrosoftUpdateService

  $filters = Prompt-SearchFilters

  $params = @{}
  $params += Get-TargetsParam
  $params += Get-VerboseParam

  if ($script:Config.UseMicrosoftUpdate) {
    $params.MicrosoftUpdate = $true
  }

  foreach ($key in $filters.Keys) {
    $params[$key] = $filters[$key]
  }

  Write-Host ""
  Write-Info "Scanning for available updates..."
  Write-Progress-Enhanced -Activity "Scanning" -Status "Querying update servers..." -PercentComplete -1

  try {
    $list = Get-WindowsUpdate @params
    $script:LastScanRaw = @($list)

    $mapped = @()
    foreach ($update in @($list)) {
      $mapped += (Get-UpdateDisplay $update)
    }
    $script:LastScan = $mapped

    Write-Progress -Activity "Scanning" -Completed

    if ($mapped.Count -eq 0) {
      Write-Host ""
      Write-Success "No updates found. System is up to date!"
      Pause
      return
    }

    Write-Host ""
    Write-Success "Found $($mapped.Count) update(s)."
    Write-Host ""

    # Display summary by severity
    $critical = @($mapped | Where-Object { $_.Severity -like "*Critical*" })
    $important = @($mapped | Where-Object { $_.Severity -like "*Important*" })
    $moderate = @($mapped | Where-Object { $_.Severity -like "*Moderate*" })
    $low = @($mapped | Where-Object { $_.Severity -like "*Low*" })

    if ($critical.Count -gt 0) {
      Write-Host "  Critical: $($critical.Count)" -ForegroundColor Red
    }
    if ($important.Count -gt 0) {
      Write-Host "  Important: $($important.Count)" -ForegroundColor Yellow
    }
    if ($moderate.Count -gt 0) {
      Write-Host "  Moderate: $($moderate.Count)" -ForegroundColor Cyan
    }
    if ($low.Count -gt 0) {
      Write-Host "  Low: $($low.Count)" -ForegroundColor Gray
    }

    Write-Host ""

    # Display table
    $mapped |
      Select-Object KB, Severity, @{N="Size (MB)";E={if($_.Size){[math]::Round($_.Size/1MB,2)}else{"N/A"}}}, Title |
      Format-Table -AutoSize |
      Out-Host

    # Export option
    if (Read-YesNo "Export scan results to CSV?") {
      $exportPath = Join-Path $env:TEMP ("WindowsUpdate_Scan_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

      try {
        $mapped |
          Select-Object KB, Severity, Size, Title, UpdateID, RevisionNumber |
          Export-Csv -NoTypeInformation -Path $exportPath -Force -ErrorAction Stop

        Write-Success "Exported to: $exportPath"

        if (Read-YesNo "Open export file?") {
          Start-Process $exportPath
        }
      } catch {
        Write-Err "Failed to export: $($_.Exception.Message)"
        Write-ErrorLog -Message "Export failed" -ErrorRecord $_
      }
    }

  } catch {
    Write-Err "Scan failed: $($_.Exception.Message)"
    Write-ErrorLog -Message "Update scan failed" -ErrorRecord $_
  } finally {
    Write-Progress -Activity "Scanning" -Completed
    Pause
  }
}

function Get-InstallCmd {
  $cmd = Get-Command Install-WindowsUpdate -ErrorAction SilentlyContinue
  if ($cmd) { return "Install-WindowsUpdate" }
  return $null
}

function Invoke-InstallByIdentity {
  param([Parameter(Mandatory)] $item)

  $params = @{}
  $params += Get-TargetsParam
  $params += Get-VerboseParam

  if ($script:Config.UseMicrosoftUpdate) {
    $params.MicrosoftUpdate = $true
  }

  $params.AcceptAll = $true

  if ($script:Config.AutoReboot) {
    $params.AutoReboot = $true
  }
  if ($script:Config.IgnoreReboot) {
    $params.IgnoreReboot = $true
  }

  $installCmd = Get-InstallCmd
  if (-not $installCmd) {
    throw "Install-WindowsUpdate cmdlet not found. Update PSWindowsUpdate or reinstall the module."
  }

  # Prefer UpdateID + RevisionNumber
  if ($item.UpdateID -and ($null -ne $item.RevisionNumber)) {
    $params.UpdateID = $item.UpdateID
    $params.RevisionNumber = $item.RevisionNumber
  } elseif ($item.KB) {
    $params.KBArticleID = $item.KB
  } else {
    $params.Title = $item.Title
  }

  & $installCmd @params
}

function Action-InstallAll {
  Write-Header "Windows Update Manager — Install ALL Updates" $script:Config

  # Pre-flight checks
  if (-not (Invoke-PreFlightChecks -Detailed)) {
    Write-Warning "Pre-flight checks failed."
    if (-not (Read-YesNo "Continue anyway?")) {
      return
    }
  }

  Ensure-MicrosoftUpdateService

  Write-Host ""
  Write-Warning "This will install ALL available updates."
  Write-Warning "This operation may take a long time and could require a reboot."
  Write-Host ""

  if (-not (Read-YesNo "Install ALL available updates now?")) {
    return
  }

  # Create restore point
  if (-not (New-SystemRestorePoint -Description "Before Windows Updates (Install All)")) {
    return
  }

  $params = @{}
  $params += Get-TargetsParam
  $params += Get-VerboseParam

  if ($script:Config.UseMicrosoftUpdate) {
    $params.MicrosoftUpdate = $true
  }

  $params.AcceptAll = $true

  if ($script:Config.AutoReboot) {
    $params.AutoReboot = $true
  }
  if ($script:Config.IgnoreReboot) {
    $params.IgnoreReboot = $true
  }

  $installCmd = Get-InstallCmd
  if (-not $installCmd) {
    Write-Err "Install-WindowsUpdate cmdlet not found. Try Update-WUModule or reinstall PSWindowsUpdate."
    Pause
    return
  }

  Write-Host ""
  Write-Warning "Starting installation... This may take a while."
  Write-Info "Progress will be displayed below."
  Write-Host ""

  try {
    & $installCmd @params | Out-Host
    Write-Host ""
    Write-Success "Install operation completed successfully!"
  } catch {
    Write-Host ""
    Write-Err "Install failed: $($_.Exception.Message)"
    Write-ErrorLog -Message "Install all failed" -ErrorRecord $_
  }

  Pause
}

function Action-InstallSelected {
  Write-Header "Windows Update Manager — Install Selected Updates" $script:Config

  if (-not $script:LastScan -or $script:LastScan.Count -eq 0) {
    Write-Warning "No cached scan results found."
    Write-Info "Please run 'Scan for updates' first (Option 2)."
    Pause
    return
  }

  $selected = Select-FromList -Items $script:LastScan -Title "Select updates to INSTALL"

  if (-not $selected -or $selected.Count -eq 0) {
    Write-Info "No updates selected."
    Pause
    return
  }

  Write-Host ""
  Write-Info "You selected $($selected.Count) update(s):"
  Write-Host ""

  $selected |
    Select-Object KB, Title |
    Format-Table -AutoSize |
    Out-Host

  # Check for excluded KBs
  $excluded = @()
  foreach ($item in $selected) {
    if ($item.KB -and ($item.KB -in $script:Config.ExcludedKBs)) {
      $excluded += $item
    }
  }

  if ($excluded.Count -gt 0) {
    Write-Warning "The following updates are in your exclusion list:"
    $excluded | Select-Object KB, Title | Format-Table -AutoSize | Out-Host

    if (-not (Read-YesNo "Install these excluded updates anyway?")) {
      return
    }
  }

  if (-not (Read-YesNo "Proceed with installation of selected updates?")) {
    return
  }

  # Create restore point
  if (-not (New-SystemRestorePoint -Description "Before Windows Updates (Selected)")) {
    return
  }

  Write-Host ""
  Write-Info "Starting installation..."
  Write-Host ""

  try {
    foreach ($item in $selected) {
      Write-Info "Installing: $($item.Display)"
      Invoke-InstallByIdentity -item $item | Out-Host
      Write-Host ""
    }

    Write-Success "Selected update installation completed!"
  } catch {
    Write-Err "Install failed: $($_.Exception.Message)"
    Write-ErrorLog -Message "Install selected failed" -ErrorRecord $_
  }

  Pause
}

function Get-HideCmd {
  $cmd = Get-Command Hide-WindowsUpdate -ErrorAction SilentlyContinue
  if ($cmd) { return "Hide-WindowsUpdate" }

  $cmd2 = Get-Command Get-WindowsUpdate -ErrorAction SilentlyContinue
  if ($cmd2) { return "Get-WindowsUpdate" }

  return $null
}

function Invoke-HideUnhide {
  param(
    [Parameter(Mandatory)]
    [ValidateSet("Hide", "Unhide")]
    [string]$Mode,

    [Parameter(Mandatory)]
    $item
  )

  $cmd = Get-HideCmd
  if (-not $cmd) {
    throw "Hide/Unhide command not available."
  }

  $params = @{}
  $params += Get-TargetsParam
  $params += Get-VerboseParam

  if ($script:Config.UseMicrosoftUpdate) {
    $params.MicrosoftUpdate = $true
  }

  if ($item.UpdateID -and ($null -ne $item.RevisionNumber)) {
    $params.UpdateID = $item.UpdateID
    $params.RevisionNumber = $item.RevisionNumber
  } elseif ($item.KB) {
    $params.KBArticleID = $item.KB
  } else {
    $params.Title = $item.Title
  }

  if ($cmd -eq "Get-WindowsUpdate") {
    if ($Mode -eq "Hide") {
      $params.Hide = $true
    }
    if ($Mode -eq "Unhide") {
      $params.Unhide = $true
    }
    & $cmd @params
  } else {
    if ($Mode -eq "Hide") {
      & $cmd @params
    } else {
      # Unhide fallback
      $params2 = @{}
      $params2 += $params
      $params2.Remove("MicrosoftUpdate") | Out-Null
      if ($script:Config.UseMicrosoftUpdate) {
        $params2.MicrosoftUpdate = $true
      }
      $params2.Unhide = $true
      Get-WindowsUpdate @params2
    }
  }
}

function Action-HideSelected {
  Write-Header "Windows Update Manager — Hide Selected Updates" $script:Config

  if (-not $script:LastScan -or $script:LastScan.Count -eq 0) {
    Write-Warning "No cached scan results found."
    Write-Info "Please run 'Scan for updates' first (Option 2)."
    Pause
    return
  }

  $selected = Select-FromList -Items $script:LastScan -Title "Select updates to HIDE"

  if (-not $selected -or $selected.Count -eq 0) {
    Write-Info "No updates selected."
    Pause
    return
  }

  Write-Host ""
  Write-Info "You selected $($selected.Count) update(s) to hide:"
  Write-Host ""

  $selected |
    Select-Object KB, Title |
    Format-Table -AutoSize |
    Out-Host

  if (-not (Read-YesNo "Hide selected updates?")) {
    return
  }

  Write-Host ""

  try {
    foreach ($item in $selected) {
      Write-Info "Hiding: $($item.Display)"
      Invoke-HideUnhide -Mode "Hide" -item $item | Out-Host

      # Add to exclusion list
      if ($item.KB -and ($item.KB -notin $script:Config.ExcludedKBs)) {
        $script:Config.ExcludedKBs += $item.KB
      }

      Write-Host ""
    }

    Save-Config $script:Config
    Write-Success "Hide operation completed and exclusion list updated!"
  } catch {
    Write-Err "Hide failed: $($_.Exception.Message)"
    Write-ErrorLog -Message "Hide operation failed" -ErrorRecord $_
  }

  Pause
}

function Action-UnhideByKB {
  Write-Header "Windows Update Manager — Unhide by KB" $script:Config

  Write-Info "Enter the KB number to unhide (e.g. KB5031234)"
  $kb = Read-Host "KB Number"

  if ([string]::IsNullOrWhiteSpace($kb)) {
    Write-Warning "No KB entered."
    Pause
    return
  }

  $kb = $kb.Trim()

  if (-not (Read-YesNo "Unhide $kb?")) {
    return
  }

  Write-Host ""

  try {
    $item = [pscustomobject]@{
      Display = $kb
      Title = $kb
      KB = $kb
      UpdateID = $null
      RevisionNumber = $null
    }

    Write-Info "Unhiding: $kb"
    Invoke-HideUnhide -Mode "Unhide" -item $item | Out-Host

    # Remove from exclusion list
    if ($kb -in $script:Config.ExcludedKBs) {
      $script:Config.ExcludedKBs = @($script:Config.ExcludedKBs | Where-Object { $_ -ne $kb })
      Save-Config $script:Config
      Write-Info "Removed from exclusion list."
    }

    Write-Host ""
    Write-Success "Unhide operation completed!"
  } catch {
    Write-Err "Unhide failed: $($_.Exception.Message)"
    Write-ErrorLog -Message "Unhide operation failed" -ErrorRecord $_
  }

  Pause
}

function Action-UninstallByKB {
  Write-Header "Windows Update Manager — Uninstall by KB" $script:Config

  Write-Info "Enter the KB number to uninstall (e.g. KB5031234)"
  $kb = Read-Host "KB Number"

  if ([string]::IsNullOrWhiteSpace($kb)) {
    Write-Warning "No KB entered."
    Pause
    return
  }

  $kb = $kb.Trim()

  $params = @{}
  $params += Get-TargetsParam
  $params += Get-VerboseParam
  $params.KBArticleID = $kb

  Write-Host ""
  Write-Warning "UNINSTALL $kb?"
  Write-Warning "This operation may require a reboot and cannot be easily undone."
  Write-Host ""

  if (-not (Read-YesNo "Proceed with uninstall?")) {
    return
  }

  # Create restore point
  if (-not (New-SystemRestorePoint -Description "Before Uninstalling $kb")) {
    return
  }

  Write-Host ""
  Write-Info "Uninstalling $kb..."
  Write-Host ""

  try {
    Remove-WindowsUpdate @params | Out-Host
    Write-Host ""
    Write-Success "Uninstall operation completed!"
  } catch {
    Write-Err "Uninstall failed: $($_.Exception.Message)"
    Write-ErrorLog -Message "Uninstall operation failed" -ErrorRecord $_
  }

  Pause
}

function Action-History {
  Write-Header "Windows Update Manager — Update History" $script:Config

  Write-Info "How many history entries to retrieve?"
  $last = Read-Host "Number of entries (blank for default)"

  $params = @{}
  $params += Get-TargetsParam
  $params += Get-VerboseParam

  if (-not [string]::IsNullOrWhiteSpace($last) -and $last -match "^\d+$") {
    $params.Last = [int]$last
  }

  Write-Host ""
  Write-Info "Retrieving update history..."
  Write-Progress-Enhanced -Activity "History" -Status "Querying update history..." -PercentComplete -1

  try {
    $history = Get-WUHistory @params
    $historyArray = @($history)

    Write-Progress -Activity "History" -Completed

    if ($historyArray.Count -eq 0) {
      Write-Host ""
      Write-Info "No history entries found."
      Pause
      return
    }

    Write-Host ""
    Write-Success "Retrieved $($historyArray.Count) history entries."
    Write-Host ""

    $historyArray |
      Select-Object Date, Result, Title, KB |
      Format-Table -AutoSize |
      Out-Host

    # Export option
    if (Read-YesNo "Export history to CSV?") {
      $exportPath = Join-Path $env:TEMP ("WindowsUpdate_History_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

      try {
        $historyArray |
          Export-Csv -NoTypeInformation -Path $exportPath -Force -ErrorAction Stop

        Write-Success "Exported to: $exportPath"

        if (Read-YesNo "Open export file?") {
          Start-Process $exportPath
        }
      } catch {
        Write-Err "Failed to export: $($_.Exception.Message)"
        Write-ErrorLog -Message "History export failed" -ErrorRecord $_
      }
    }

  } catch {
    Write-Err "History retrieval failed: $($_.Exception.Message)"
    Write-ErrorLog -Message "History retrieval failed" -ErrorRecord $_
    Write-Info "Note: Some environments have long history. If it hangs, limit with a number."
  } finally {
    Write-Progress -Activity "History" -Completed
    Pause
  }
}

function Action-ResetWUComponents {
  Write-Header "Windows Update Manager — Reset Components" $script:Config

  Write-Host ""
  Write-Warning "CAUTION: This will reset Windows Update components."
  Write-Warning "This operation will:"
  Write-Host "  • Stop Windows Update services" -ForegroundColor Yellow
  Write-Host "  • Clear update cache and temporary files" -ForegroundColor Yellow
  Write-Host "  • Re-register Windows Update components" -ForegroundColor Yellow
  Write-Host "  • May require a system reboot" -ForegroundColor Yellow
  Write-Host ""
  Write-Info "Use this only if Windows Update is experiencing persistent problems."
  Write-Host ""

  if (-not (Read-YesNo "Proceed with Reset-WUComponents?")) {
    return
  }

  # Create restore point
  if (-not (New-SystemRestorePoint -Description "Before Windows Update Component Reset")) {
    return
  }

  Write-Host ""
  Write-Info "Resetting Windows Update components..."
  Write-Info "This may take several minutes..."
  Write-Host ""

  try {
    Reset-WUComponents @(Get-TargetsParam) @(Get-VerboseParam) | Out-Host
    Write-Host ""
    Write-Success "Windows Update components reset successfully!"
    Write-Info "A reboot may be required for changes to take effect."
  } catch {
    Write-Err "Reset failed: $($_.Exception.Message)"
    Write-ErrorLog -Message "Component reset failed" -ErrorRecord $_
  }

  Pause
}

function Action-UpdateModule {
  Write-Header "Windows Update Manager — Update PSWindowsUpdate Module" $script:Config

  Write-Info "Current PSWindowsUpdate version:"
  $currentModule = Get-Module -ListAvailable PSWindowsUpdate | Sort-Object Version -Descending | Select-Object -First 1

  if ($currentModule) {
    Write-Host "  Version: $($currentModule.Version)"
  } else {
    Write-Warning "PSWindowsUpdate module not found."
  }

  Write-Host ""
  Write-Info "This will check for and install the latest version from PowerShell Gallery."
  Write-Host ""

  if (-not (Read-YesNo "Update PSWindowsUpdate module now?")) {
    return
  }

  Write-Host ""
  Write-Info "Updating PSWindowsUpdate module..."

  try {
    Update-WUModule @(Get-VerboseParam) | Out-Host
    Write-Host ""
    Write-Success "PSWindowsUpdate module update completed!"
    Write-Info "You may need to restart this script to use the new version."
  } catch {
    Write-Err "Module update failed: $($_.Exception.Message)"
    Write-ErrorLog -Message "Module update failed" -ErrorRecord $_
  }

  Pause
}

function Action-Targets {
  Write-Header "Windows Update Manager — Target Computers" $script:Config

  Write-Info "Configure target computers for remote update management."
  Write-Info "Leave blank to manage the local computer only."
  Write-Host ""
  Write-Info "Current targets:"

  if ($script:Config.TargetComputers.Count -gt 0) {
    foreach ($target in $script:Config.TargetComputers) {
      Write-Host "  • $target"
    }
  } else {
    Write-Host "  • LOCAL (this computer)"
  }

  Write-Host ""
  Write-Info "Enter new target computers (comma-separated) or blank for local only:"
  $raw = Read-Host "Targets"

  if ([string]::IsNullOrWhiteSpace($raw)) {
    $script:Config.TargetComputers = @()
  } else {
    $names = @()
    foreach ($name in $raw.Split(",") | ForEach-Object { $_.Trim() } | Where-Object { $_ }) {
      $names += $name
    }
    $script:Config.TargetComputers = @($names)
  }

  Save-Config $script:Config

  Write-Host ""
  Write-Success "Target computers updated!"

  if ($script:Config.TargetComputers.Count -gt 0) {
    Write-Info "New targets:"
    foreach ($target in $script:Config.TargetComputers) {
      Write-Host "  • $target"
    }
  } else {
    Write-Info "Targeting local computer only."
  }

  Pause
}

function Action-Settings {
  while ($true) {
    Write-Header "Windows Update Manager — Settings" $script:Config

    Write-Host "  [1] Toggle Update Source (Windows Update / Microsoft Update)"
    Write-Host "      Switch the update catalog used for scans and installs"
    Write-Host "      Current: $(if ($script:Config.UseMicrosoftUpdate) { 'Microsoft Update' } else { 'Windows Update' })"
    Write-Host ""
    Write-Host "  [2] Toggle Auto Reboot"
    Write-Host "      Reboot automatically after updates when required"
    Write-Host "      Current: $(if ($script:Config.AutoReboot) { 'Enabled' } else { 'Disabled' })"
    Write-Host ""
    Write-Host "  [3] Toggle Ignore Reboot"
    Write-Host "      Ignore reboot-required checks during operations"
    Write-Host "      Current: $(if ($script:Config.IgnoreReboot) { 'Enabled' } else { 'Disabled' })"
    Write-Host ""
    Write-Host "  [4] Toggle Verbose Output"
    Write-Host "      Show detailed module output and diagnostics"
    Write-Host "      Current: $(if ($script:Config.Verbose) { 'Enabled' } else { 'Disabled' })"
    Write-Host ""
    Write-Host "  [5] Toggle System Restore Point Creation"
    Write-Host "      Create a restore point before update actions"
    Write-Host "      Current: $(if ($script:Config.CreateRestorePoint) { 'Enabled' } else { 'Disabled' })"
    Write-Host ""
    Write-Host "  [6] Toggle Logging (Transcript)"
    Write-Host "      Start or stop transcript logging to file"
    Write-Host "      Current: $(if ($script:TranscriptOn) { 'ON' } else { 'OFF' })"
    Write-Host ""
    Write-Host "  [7] Manage KB Exclusion List"
    Write-Host "      Add or remove KBs to flag during installs"
    Write-Host "      Current: $($script:Config.ExcludedKBs.Count) excluded KB(s)"
    Write-Host ""
    Write-Host "  [0] Back to Main Menu"
    Write-Host "      Return to the main menu"
    Write-Host ""

    $choice = Read-Host "Choice"

    switch ($choice) {
      "1" {
        $script:Config.UseMicrosoftUpdate = -not $script:Config.UseMicrosoftUpdate
        Save-Config $script:Config
        Write-Success "Update source toggled!"
        Start-Sleep -Seconds 1
      }
      "2" {
        $script:Config.AutoReboot = -not $script:Config.AutoReboot
        Save-Config $script:Config
        Write-Success "Auto reboot toggled!"
        Start-Sleep -Seconds 1
      }
      "3" {
        $script:Config.IgnoreReboot = -not $script:Config.IgnoreReboot
        Save-Config $script:Config
        Write-Success "Ignore reboot toggled!"
        Start-Sleep -Seconds 1
      }
      "4" {
        $script:Config.Verbose = -not $script:Config.Verbose
        Save-Config $script:Config
        Write-Success "Verbose output toggled!"
        Start-Sleep -Seconds 1
      }
      "5" {
        $script:Config.CreateRestorePoint = -not $script:Config.CreateRestorePoint
        Save-Config $script:Config
        Write-Success "Restore point creation toggled!"
        Start-Sleep -Seconds 1
      }
      "6" {
        Toggle-Transcript
      }
      "7" {
        Manage-ExclusionList
      }
      "0" {
        return
      }
      default { }
    }
  }
}

function Toggle-Transcript {
  if ($script:TranscriptOn) {
    try {
      Stop-Transcript | Out-Null
    } catch { }

    $script:TranscriptOn = $false
    Write-Success "Logging disabled."
    Pause
    return
  }

  $path = $LogPath

  if ([string]::IsNullOrWhiteSpace($path)) {
    $logDir = Join-Path $env:APPDATA "WindowsUpdateManager\logs"
    if (-not (Test-Path $logDir)) {
      New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    $path = Join-Path $logDir ("Transcript_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
  }

  try {
    Start-Transcript -Path $path -Append | Out-Null
    $script:TranscriptOn = $true
    Write-Success "Logging enabled: $path"
  } catch {
    Write-Err "Failed to start transcript: $($_.Exception.Message)"
    Write-ErrorLog -Message "Transcript start failed" -ErrorRecord $_
  }

  Pause
}

function Manage-ExclusionList {
  while ($true) {
    Write-Header "Windows Update Manager — KB Exclusion List" $script:Config

    Write-Info "Updates in this list will be flagged during installation."
    Write-Host ""

    if ($script:Config.ExcludedKBs.Count -eq 0) {
      Write-Info "Exclusion list is empty."
    } else {
      Write-Info "Excluded KBs:"
      foreach ($kb in $script:Config.ExcludedKBs) {
        Write-Host "  • $kb"
      }
    }

    Write-Host ""
    Write-Host "  [1] Add KB to exclusion list"
    Write-Host "      Add a KB to flag during installs"
    Write-Host "  [2] Remove KB from exclusion list"
    Write-Host "      Remove a KB from the list"
    Write-Host "  [3] Clear all exclusions"
    Write-Host "      Remove all excluded KB entries"
    Write-Host "  [0] Back"
    Write-Host "      Return to Settings"
    Write-Host ""

    $choice = Read-Host "Choice"

    switch ($choice) {
      "1" {
        $kb = Read-Host "Enter KB to exclude (e.g. KB5031234)"
        if (-not [string]::IsNullOrWhiteSpace($kb)) {
          $kb = $kb.Trim()
          if ($kb -notin $script:Config.ExcludedKBs) {
            $script:Config.ExcludedKBs += $kb
            Save-Config $script:Config
            Write-Success "$kb added to exclusion list."
          } else {
            Write-Warning "$kb is already in the exclusion list."
          }
        }
        Start-Sleep -Seconds 2
      }
      "2" {
        if ($script:Config.ExcludedKBs.Count -eq 0) {
          Write-Warning "Exclusion list is empty."
          Start-Sleep -Seconds 2
          continue
        }

        $kb = Read-Host "Enter KB to remove"
        if (-not [string]::IsNullOrWhiteSpace($kb)) {
          $kb = $kb.Trim()
          if ($kb -in $script:Config.ExcludedKBs) {
            $script:Config.ExcludedKBs = @($script:Config.ExcludedKBs | Where-Object { $_ -ne $kb })
            Save-Config $script:Config
            Write-Success "$kb removed from exclusion list."
          } else {
            Write-Warning "$kb is not in the exclusion list."
          }
        }
        Start-Sleep -Seconds 2
      }
      "3" {
        if (Read-YesNo "Clear all KB exclusions?") {
          $script:Config.ExcludedKBs = @()
          Save-Config $script:Config
          Write-Success "Exclusion list cleared."
          Start-Sleep -Seconds 2
        }
      }
      "0" {
        return
      }
      default { }
    }
  }
}

function Action-RemoteInstallJob {
  Write-Header "Windows Update Manager — Remote Install Job" $script:Config

  Write-Info "This uses Invoke-WUJob to create a Scheduled Task running as SYSTEM on remote computers."
  Write-Info "This provides reliable remote installations."
  Write-Host ""

  Write-Info "Enter target computers (comma-separated):"
  $raw = Read-Host "Targets"

  if ([string]::IsNullOrWhiteSpace($raw)) {
    Write-Warning "No targets specified."
    Pause
    return
  }

  $targets = @()
  foreach ($name in $raw.Split(",") | ForEach-Object { $_.Trim() } | Where-Object { $_ }) {
    $targets += $name
  }

  Write-Host ""
  $useMU = Read-YesNo "Use Microsoft Update?" -DefaultYes:$script:Config.UseMicrosoftUpdate
  $autoReboot = Read-YesNo "Auto reboot if needed?" -DefaultYes:$script:Config.AutoReboot

  $scriptBlock = {
    Import-Module PSWindowsUpdate -ErrorAction Stop
    $params = @{ AcceptAll = $true }
    if ($using:useMU) { $params.MicrosoftUpdate = $true }
    if ($using:autoReboot) { $params.AutoReboot = $true }
    Install-WindowsUpdate @params | Out-File -FilePath "$env:TEMP\WindowsUpdate_RemoteJob.log" -Append
  }

  Write-Host ""
  Write-Info "Target computers: $($targets -join ', ')"
  Write-Info "Microsoft Update: $useMU"
  Write-Info "Auto Reboot: $autoReboot"
  Write-Host ""

  if (-not (Read-YesNo "Create SYSTEM install job on specified computers?")) {
    return
  }

  Write-Host ""
  Write-Info "Submitting remote job..."

  try {
    $cmd = Get-Command Invoke-WUJob -ErrorAction Stop

    $params = @{
      ComputerName = $targets
      Confirm = $false
    }

    if ($cmd.Parameters.ContainsKey("RunNow")) {
      $params.RunNow = $true
    }

    if ($cmd.Parameters.ContainsKey("Script")) {
      $params.Script = $scriptBlock
    } elseif ($cmd.Parameters.ContainsKey("ScriptBlock")) {
      $params.ScriptBlock = $scriptBlock
    } else {
      throw "Invoke-WUJob has no -Script/-ScriptBlock parameter in this environment."
    }

    Invoke-WUJob @params | Out-Host

    Write-Host ""
    Write-Success "Remote job submitted successfully!"
    Write-Info "Check remote log: %TEMP%\WindowsUpdate_RemoteJob.log"
  } catch {
    Write-Err "Remote job failed: $($_.Exception.Message)"
    Write-ErrorLog -Message "Remote job failed" -ErrorRecord $_
    Write-Info "Ensure PowerShell remoting is enabled and PSWindowsUpdate is available on targets."
    Write-Info "You can use Enable-WURemoting on remote computers."
  }

  Pause
}

function Action-ComplianceReport {
  Write-Header "Windows Update Manager — Compliance Report" $script:Config

  Write-Info "Generating compliance report..."
  Write-Host ""

  $reportPath = Join-Path $env:TEMP ("WindowsUpdate_Compliance_{0}.txt" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

  try {
    $report = @()
    $report += "=" * 80
    $report += "Windows Update Compliance Report"
    $report += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $report += "=" * 80
    $report += ""

    # System info
    $report += "SYSTEM INFORMATION"
    $report += "-" * 80
    $report += "Computer: $env:COMPUTERNAME"
    $report += "OS: $(Get-CimInstance Win32_OperatingSystem | Select-Object -ExpandProperty Caption)"
    $report += "PS Version: $($PSVersionTable.PSVersion)"
    $report += ""

    # Update status
    $report += "UPDATE STATUS"
    $report += "-" * 80

    try {
      $status = Get-WUInstallerStatus -ErrorAction Stop
      $report += "Installer Status: $($status.Status)"
    } catch {
      $report += "Installer Status: Unable to retrieve"
    }

    try {
      $rebootStatus = Get-WURebootStatus -ErrorAction Stop
      $report += "Reboot Required: $($rebootStatus.RebootRequired)"
    } catch {
      $report += "Reboot Required: Unable to determine"
    }

    $report += ""

    # Available updates
    $report += "AVAILABLE UPDATES"
    $report += "-" * 80

    try {
      $updates = Get-WindowsUpdate -MicrosoftUpdate
      if ($updates) {
        $updateArray = @($updates)
        $report += "Total Available: $($updateArray.Count)"
        $report += ""

        foreach ($update in $updateArray) {
          $kb = if ($update.KB) { $update.KB } else { "N/A" }
          $report += "  [$kb] $($update.Title)"
        }
      } else {
        $report += "No updates available. System is up to date."
      }
    } catch {
      $report += "Unable to retrieve available updates."
    }

    $report += ""
    $report += "=" * 80
    $report += "End of Report"
    $report += "=" * 80

    # Save report
    $report | Out-File -FilePath $reportPath -Encoding UTF8 -Force

    Write-Success "Compliance report generated: $reportPath"
    Write-Host ""

    if (Read-YesNo "Open report file?") {
      Start-Process notepad.exe -ArgumentList $reportPath
    }

  } catch {
    Write-Err "Failed to generate compliance report: $($_.Exception.Message)"
    Write-ErrorLog -Message "Compliance report failed" -ErrorRecord $_
  }

  Pause
}

#endregion

#region Main Menu

function Show-MainMenu {
  $menuLines = @(
    "UPDATE OPERATIONS:",
    "  [1]  Status",
    "      View current update system status and health",
    "  [2]  Scan for updates",
    "      Check for available Windows updates",
    "  [3]  Install ALL updates",
    "      Download and install all available updates",
    "  [4]  Install SELECTED updates",
    "      Choose which updates to install from scan",
    "",
    "UPDATE MANAGEMENT:",
    "  [5]  Hide SELECTED updates",
    "      Prevent unwanted updates from appearing",
    "  [6]  Unhide by KB",
    "      Make previously hidden updates visible again",
    "  [7]  Uninstall by KB",
    "      Remove a specific update from the system",
    "  [8]  Show update history",
    "      View past installations and export to CSV",
    "",
    "SYSTEM MAINTENANCE:",
    "  [9]  Reset Windows Update components",
    "      Fix update problems",
    "  [10] Update PSWindowsUpdate module",
    "      Get latest version of update tool",
    "",
    "CONFIGURATION:",
    "  [11] Target computers",
    "      Manage local or remote systems",
    "  [12] Settings",
    "      Configure update behavior and visuals",
    "",
    "ADVANCED:",
    "  [13] Remote install job",
    "      Schedule updates on remote computers",
    "  [14] Generate compliance report",
    "      Create audit report of update status",
    "",
    "  [0]  Exit",
    "      Quit Windows Update Manager",
    ""
  )

  $maxLine = ($menuLines | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum
  $requiredWidth = $maxLine + 2
  $requiredHeight = 10 + $menuLines.Count + 2

  # Auto-size and center console window
  Set-OptimalWindowSize -RequiredWidth $requiredWidth -RequiredHeight $requiredHeight

  Write-Header "Windows Update Manager — Main Menu" $script:Config

  Write-Host "UPDATE OPERATIONS:" -ForegroundColor Cyan
  Write-Host "  [1]  Status" -ForegroundColor White
  Write-Host "      View current update system status and health" -ForegroundColor Gray
  Write-Host "  [2]  Scan for updates" -ForegroundColor White
  Write-Host "      Check for available Windows updates" -ForegroundColor Gray
  Write-Host "  [3]  Install ALL updates" -ForegroundColor White
  Write-Host "      Download and install all available updates" -ForegroundColor Gray
  Write-Host "  [4]  Install SELECTED updates" -ForegroundColor White
  Write-Host "      Choose which updates to install from scan" -ForegroundColor Gray
  Write-Host ""

  Write-Host "UPDATE MANAGEMENT:" -ForegroundColor Cyan
  Write-Host "  [5]  Hide SELECTED updates" -ForegroundColor White
  Write-Host "      Prevent unwanted updates from appearing" -ForegroundColor Gray
  Write-Host "  [6]  Unhide by KB" -ForegroundColor White
  Write-Host "      Make previously hidden updates visible again" -ForegroundColor Gray
  Write-Host "  [7]  Uninstall by KB" -ForegroundColor White
  Write-Host "      Remove a specific update from the system" -ForegroundColor Gray
  Write-Host "  [8]  Show update history" -ForegroundColor White
  Write-Host "      View past installations and export to CSV" -ForegroundColor Gray
  Write-Host ""

  Write-Host "SYSTEM MAINTENANCE:" -ForegroundColor Cyan
  Write-Host "  [9]  Reset Windows Update components" -ForegroundColor White
  Write-Host "      Fix update problems" -ForegroundColor Gray
  Write-Host "  [10] Update PSWindowsUpdate module" -ForegroundColor White
  Write-Host "      Get latest version of update tool" -ForegroundColor Gray
  Write-Host ""

  Write-Host "CONFIGURATION:" -ForegroundColor Cyan
  Write-Host "  [11] Target computers" -ForegroundColor White
  Write-Host "      Manage local or remote systems" -ForegroundColor Gray
  Write-Host "  [12] Settings" -ForegroundColor White
  Write-Host "      Configure update behavior and visuals" -ForegroundColor Gray
  Write-Host ""

  Write-Host "ADVANCED:" -ForegroundColor Cyan
  Write-Host "  [13] Remote install job" -ForegroundColor White
  Write-Host "      Schedule updates on remote computers" -ForegroundColor Gray
  Write-Host "  [14] Generate compliance report" -ForegroundColor White
  Write-Host "      Create audit report of update status" -ForegroundColor Gray
  Write-Host ""

  Write-Host "  [0]  Exit" -ForegroundColor Yellow
  Write-Host "      Quit Windows Update Manager" -ForegroundColor Gray
  Write-Host ""

  return (Read-Host "Choice")
}

#endregion

#region Entry Point

function Main {
  try {
    # Initialize
    Initialize-Colors
    Initialize-ErrorLog

    # Show banner
    Show-Banner

    Write-Info "Initializing Windows Update Manager v$($script:Version)..."
    Write-Host ""

    # Check elevation
    Ensure-Elevation -Skip:$NoElevation
    $script:IsElevated = Test-IsAdmin

    # Initialize dependencies
    if (-not (Initialize-Dependencies)) {
      Write-Err "Failed to initialize dependencies. Cannot continue."
      Write-Info "Error log: $($script:ErrorLogPath)"
      Pause
      exit 1
    }

    # Load configuration
    $script:Config = Load-Config

    # Check if first run
    if (-not (Test-Path $ConfigPath)) {
      Show-FirstRunWizard
    }

    # Main loop
    while ($true) {
      $choice = Show-MainMenu

      switch ($choice) {
        "1"  { Action-ShowStatus }
        "2"  { Action-ScanUpdates }
        "3"  { Action-InstallAll }
        "4"  { Action-InstallSelected }
        "5"  { Action-HideSelected }
        "6"  { Action-UnhideByKB }
        "7"  { Action-UninstallByKB }
        "8"  { Action-History }
        "9"  { Action-ResetWUComponents }
        "10" { Action-UpdateModule }
        "11" { Action-Targets }
        "12" { Action-Settings }
        "13" { Action-RemoteInstallJob }
        "14" { Action-ComplianceReport }
        "0"  {
          Write-Info "Exiting Windows Update Manager..."
          break
        }
        default { }
      }

      if ($choice -eq "0") { break }
    }

  } catch {
    Write-Err "Fatal error: $($_.Exception.Message)"
    Write-ErrorLog -Message "Fatal error" -ErrorRecord $_
    Pause
    exit 1
  } finally {
    # Cleanup
    if ($script:TranscriptOn) {
      try {
        Stop-Transcript | Out-Null
      } catch { }
    }

    if ($script:ErrorLogPath -and (Test-Path $script:ErrorLogPath)) {
      Write-Host ""
      Write-Info "Error log: $($script:ErrorLogPath)"
    }
  }
}

# Run main
Main

# Additional helper functions needed for v3

function Read-YesNo {
    param(
        [Parameter(Mandatory)]
        [string]$Prompt,
        [switch]$DefaultYes
    )
    
    if ($Silent) { return [bool]$DefaultYes }
    
    $suffix = if ($DefaultYes) { " [Y/n]" } else { " [y/N]" }
    $answer = Read-Host ($Prompt + $suffix)
    
    if ([string]::IsNullOrWhiteSpace($answer)) {
        return [bool]$DefaultYes
    }
    
    return ($answer.Trim().ToLowerInvariant() -in @("y","yes"))
}

function Pause {
    Write-Host ""
    Write-Host "Press any key to continue..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Ensure-MicrosoftUpdateService {
    if ($script:Config.UseMicrosoftUpdate) {
        try {
            Add-WUServiceManager -MicrosoftUpdate -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        } catch {
            # Silently continue if already registered
        }
    }
}


#endregion

# Entry Point - Execute Main
Main

