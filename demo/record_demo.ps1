# Записує відео демо: відкриває окреме вікно консолі, запускає StudyHelper і пише це вікно через ffmpeg.
# Запуск з кореня репозиторію:  powershell -ExecutionPolicy Bypass -File demo\record_demo.ps1
# Результат: demo\demo.mp4 (вставити в слайд 6 презентації).
param([string]$Sample = "samples\conspect2_cloud.txt", [int]$Questions = 4)

$root = Split-Path $PSScriptRoot -Parent
$out = Join-Path $PSScriptRoot "demo.mp4"
$raw = Join-Path $env:TEMP "studyhelper_demo.mkv"
$title = "StudyHelper demo"

dotnet build (Join-Path $root "StudyHelper") -v q -nologo | Out-Null

$inner = @"
mode con: cols=100 lines=45
`$Host.UI.RawUI.WindowTitle = '$title'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
Set-Location '$root'
Clear-Host
Start-Sleep 3
function Show-Command(`$text) {
    Write-Host -NoNewline 'PS> ' -ForegroundColor DarkGray
    foreach (`$c in `$text.ToCharArray()) { Write-Host -NoNewline `$c -ForegroundColor Yellow; Start-Sleep -Milliseconds 35 }
    Write-Host ''
}
Show-Command 'Get-Content $Sample'
Get-Content '$Sample' -Encoding utf8
Write-Host ''
Start-Sleep 5
Show-Command 'dotnet run --project StudyHelper -- $Sample $Questions'
dotnet run --no-build --project StudyHelper -- '$Sample' $Questions
Start-Sleep 12
"@
$script = Join-Path $env:TEMP "studyhelper_demo_inner.ps1"
Set-Content -Path $script -Value $inner -Encoding utf8

$console = Start-Process conhost.exe -ArgumentList "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$script`"" -PassThru
Start-Sleep 2
$ffmpeg = Start-Process ffmpeg -PassThru -NoNewWindow -RedirectStandardError (Join-Path $env:TEMP "studyhelper_ffmpeg.log") `
    -ArgumentList "-y -loglevel error -f gdigrab -draw_mouse 0 -framerate 15 -i title=`"$title`" -c:v libx264 -preset veryfast -pix_fmt yuv420p `"$raw`""
$console.WaitForExit()
Start-Sleep 2
if (-not $ffmpeg.HasExited) { Stop-Process $ffmpeg.Id }

# Ширину і висоту вирівнюємо до парних, інакше H.264 не збереться
ffmpeg -y -loglevel error -i $raw -vf "scale=trunc(iw/2)*2:trunc(ih/2)*2" -c:v libx264 -pix_fmt yuv420p -movflags +faststart $out
Write-Host "Готово: $out"
