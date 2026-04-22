# Script para exportar slides selecionados dos arquivos PPTX
# Execute com: powershell -ExecutionPolicy Bypass -File exportar_slides.ps1

$portfolioPath = "C:\Users\Axel\Documents\Cris\Programacao\Portfolio"
$imgPath = "$portfolioPath\img"

if (-not (Test-Path $imgPath)) { New-Item -ItemType Directory -Path $imgPath | Out-Null }

# Slides a exportar por arquivo
$projetos = @(
  @{
    arquivo = "ativações ftd travessia_1303.pptx"
    slides  = @(1, 3, 4, 5, 6, 7, 8, 9)
    prefix  = "ftd"
  },
  @{
    arquivo = "Zurich_bancocarrefour_0603.pptx"
    slides  = @(1, 2, 3, 4, 5, 9, 14, 15, 17)
    prefix  = "zurich"
  },
  @{
    arquivo = "Mastercard_EG26_v2.pptx"
    slides  = @(1, 4, 5, 6, 8, 9)
    prefix  = "mastercard"
  },
  @{
    arquivo = "Votorantim_campanhadeincentivo.pptx"
    slides  = @(2, 3, 5, 12)
    prefix  = "votorantim"
  },
  @{
    arquivo = "Amigoz_campanhadeincentivo_v1.pptx"
    slides  = @(2, 3, 5, 8, 15, 16)
    prefix  = "amigoz"
  },
  @{
    arquivo = "INTEGRA_ON_v1.pptx"
    slides  = @(1, 2, 3, 5, 10)
    prefix  = "integra"
  }
)

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

foreach ($proj in $projetos) {
  $filePath = Join-Path $portfolioPath $proj.arquivo
  if (-not (Test-Path $filePath)) {
    Write-Host "ARQUIVO NAO ENCONTRADO: $($proj.arquivo)" -ForegroundColor Red
    continue
  }

  Write-Host "Abrindo: $($proj.arquivo)" -ForegroundColor Cyan
  $pres = $ppt.Presentations.Open($filePath, $true, $false, $false)
  $total = $pres.Slides.Count
  Write-Host "  Total de slides: $total"

  foreach ($num in $proj.slides) {
    if ($num -gt $total) {
      Write-Host "  Slide $num nao existe (total: $total)" -ForegroundColor Yellow
      continue
    }
    $outFile = Join-Path $imgPath "$($proj.prefix)-slide$('{0:D2}' -f $num).png"
    $pres.Slides($num).Export($outFile, "PNG", 1920, 1080)
    Write-Host "  Exportado: $outFile" -ForegroundColor Green
  }

  $pres.Close()
}

$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null

Write-Host "`nConcluido! Imagens salvas em: $imgPath" -ForegroundColor Green
