param(
  [Parameter(Mandatory = $true)]
  [string]$WebAppUrl
)

$ErrorActionPreference = 'Stop'
$WebAppUrl = $WebAppUrl.Trim()

if ($WebAppUrl -notmatch '^https://script\.google\.com/macros/s/[A-Za-z0-9_-]+/exec(?:\?.*)?$') {
  throw 'A URL parece inválida. Use a URL de implantação terminada em /exec.'
}

$folder = Split-Path -Parent $MyInvocation.MyCommand.Path
$files = @('index.html', 'organizador.html')
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$pattern = '(<meta\s+name=["'']gas-webapp-url["'']\s+content=["''])[^"'']*(["'']\s*>)'

foreach ($fileName in $files) {
  $path = Join-Path $folder $fileName
  if (-not (Test-Path -LiteralPath $path)) {
    throw "Arquivo não encontrado: $fileName"
  }

  $content = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
  if ($content -notmatch $pattern) {
    throw "Marcador gas-webapp-url não encontrado em $fileName"
  }

  [System.IO.File]::WriteAllText("$path.backup", $content, $utf8NoBom)
  $replacement = '${1}' + $WebAppUrl + '${2}'
  $updated = [System.Text.RegularExpressions.Regex]::Replace($content, $pattern, $replacement, 1)
  [System.IO.File]::WriteAllText($path, $updated, $utf8NoBom)
}

Write-Host ''
Write-Host 'URL configurada com sucesso em index.html e organizador.html.' -ForegroundColor Green
Write-Host 'Arquivos .backup foram criados antes da alteração.' -ForegroundColor DarkGray
