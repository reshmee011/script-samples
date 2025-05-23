#!/usr/bin/env pwsh

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

# Data Generated Content
./main/docfx/report-matrix.ps1 -BaseDir "./main" -ReportFile "./main/docfx/matrix.md"
./main/docfx/report-metadata.ps1 -BaseDir "./main" -ReportFile "./main/docfx/metadata.md"
./main/docfx/report-cmdusage.ps1 -BaseDir "./main" -ReportFile "./main/docfx/cmdusage.md"
./main/docfx/report-age.ps1 -BaseDir "./main" -ReportFile "./main/docfx/age.md"
./main/docfx/generate-samplesJson.ps1 -BaseDir "./main" -OutputFile "./main/docfx/samples.json"

# Main DocFX build
docfx build ./main/docfx/docfx.json --warningsAsErrors $args

# Copy the created site to the pnpcoredocs folder (= clone of the gh-pages branch)
#Remove-Item ./gh-pages/*/* -Recurse -Force
copy-item -Force -Recurse ./main/docfx/_site/* -Destination "./gh-pages"
