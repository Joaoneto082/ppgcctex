# Compila o ppgcctex no Windows (equivalente ao Makefile, sem precisar de 'make').
# Uso: na pasta do projeto, execute:  .\compile.ps1

$ErrorActionPreference = 'Continue'
Set-Location $PSScriptRoot

$tex = 'documento.tex'
$base = 'documento'

Write-Host '1/6 pdflatex (1ª passagem)...' -ForegroundColor Cyan
pdflatex -interaction=nonstopmode $tex
if ($LASTEXITCODE -ne 0) { Write-Host 'pdflatex falhou. Veja documento.log' -ForegroundColor Red; exit $LASTEXITCODE }

Write-Host '2/6 bibtex...' -ForegroundColor Cyan
bibtex $base

Write-Host '3/6 makeglossaries...' -ForegroundColor Cyan
makeglossaries $base

Write-Host '4/6 makeindex...' -ForegroundColor Cyan
makeindex $base

Write-Host '5/6 pdflatex (2ª passagem)...' -ForegroundColor Cyan
pdflatex -interaction=nonstopmode $tex

Write-Host '6/6 pdflatex (3ª passagem)...' -ForegroundColor Cyan
pdflatex -interaction=nonstopmode $tex

if (Test-Path 'documento.pdf') {
    Write-Host 'Pronto: documento.pdf gerado.' -ForegroundColor Green
} else {
    Write-Host 'PDF nao encontrado. Abra documento.log e procure por lines com !' -ForegroundColor Red
    exit 1
}
