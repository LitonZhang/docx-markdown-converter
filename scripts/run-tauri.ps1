param(
  [Parameter(Mandatory=$true)]
  [ValidateSet("dev", "build")]
  [string]$Mode
)

$vsCandidates = @(
  "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\Common7\Tools\VsDevCmd.bat",
  "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\Tools\VsDevCmd.bat"
)

$vsDevCmd = $vsCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $vsDevCmd) {
  throw "VsDevCmd.bat not found. Please install Visual Studio Build Tools with C++ workload."
}

$tauriCommand = if ($Mode -eq "dev") { "npx tauri dev" } else { "npx tauri build" }
$cmdLine = '"{0}" -arch=amd64 -host_arch=amd64 >nul && set PATH=%USERPROFILE%\\.cargo\\bin;%PATH% && {1}' -f $vsDevCmd, $tauriCommand

Write-Host "Using VS environment: $vsDevCmd"
Write-Host "Running: $tauriCommand"

cmd /c $cmdLine
if ($LASTEXITCODE -ne 0) {
  exit $LASTEXITCODE
}
