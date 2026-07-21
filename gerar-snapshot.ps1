# Script PowerShell para gerar snapshot diário da homologação
# Lê o CAF.xlsx e salva o estado atual em historico.json

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$xlsxPath = Join-Path $scriptDir "CAF.xlsx"
$historicoPath = Join-Path $scriptDir "historico.json"

# Verificar se a planilha existe
if (-not (Test-Path $xlsxPath)) {
    Write-Host "ERRO: Arquivo CAF.xlsx nao encontrado!" -ForegroundColor Red
    exit 1
}

# Abrir Excel via COM
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $workbook = $excel.Workbooks.Open($xlsxPath, 0, $true) # ReadOnly
    
    # Encontrar a aba de cenários (primeira aba ou a que contém "cen")
    $sheet = $null
    foreach ($ws in $workbook.Worksheets) {
        if ($ws.Name -match "cen|homologa") {
            $sheet = $ws
            break
        }
    }
    if ($null -eq $sheet) { $sheet = $workbook.Worksheets.Item(1) }
    
    # Ler headers da primeira linha
    $lastCol = $sheet.UsedRange.Columns.Count
    $lastRow = $sheet.UsedRange.Rows.Count
    
    $headers = @{}
    for ($col = 1; $col -le $lastCol; $col++) {
        $val = $sheet.Cells.Item(1, $col).Text
        if ($val) { $headers[$val.Trim()] = $col }
    }
    
    # Encontrar colunas relevantes
    $colNumero = $headers["Número do Cenário"]
    if (-not $colNumero) { $colNumero = $headers["Numero do Cenario"] }
    if (-not $colNumero) { $colNumero = $headers["Cenário"] }
    
    $colStatus = $headers["Status"]
    $colHomologador = $headers["Homologador"]
    
    if (-not $colNumero) {
        Write-Host "ERRO: Coluna de numero do cenario nao encontrada!" -ForegroundColor Red
        $workbook.Close($false)
        $excel.Quit()
        exit 1
    }
    
    # Contar status dos cenários
    $statusCounts = @{}
    $total = 0
    $homologadores = @{}
    
    for ($row = 2; $row -le $lastRow; $row++) {
        $numero = $sheet.Cells.Item($row, $colNumero).Text
        if (-not $numero -or -not $numero.StartsWith("CT-")) { continue }
        
        $total++
        $status = ""
        if ($colStatus) { $status = $sheet.Cells.Item($row, $colStatus).Text.Trim() }
        if (-not $status) { $status = "Não Iniciado" }
        
        if ($statusCounts.ContainsKey($status)) {
            $statusCounts[$status]++
        } else {
            $statusCounts[$status] = 1
        }
        
        # Contar por homologador
        if ($colHomologador) {
            $homolog = $sheet.Cells.Item($row, $colHomologador).Text.Trim()
            if ($homolog) {
                if (-not $homologadores.ContainsKey($homolog)) {
                    $homologadores[$homolog] = @{}
                }
                if ($homologadores[$homolog].ContainsKey($status)) {
                    $homologadores[$homolog][$status]++
                } else {
                    $homologadores[$homolog][$status] = 1
                }
            }
        }
    }
    
    $workbook.Close($false)
    $excel.Quit()
    
    # Montar snapshot do dia
    $hoje = Get-Date -Format "yyyy-MM-dd"
    $horaAtual = Get-Date -Format "HH:mm"
    
    $snapshot = @{
        data = $hoje
        hora = $horaAtual
        total = $total
        status = $statusCounts
        homologadores = $homologadores
    }
    
    # Ler histórico existente ou criar novo
    $historico = @()
    if (Test-Path $historicoPath) {
        $content = Get-Content $historicoPath -Raw -Encoding UTF8
        if ($content) {
            $historico = $content | ConvertFrom-Json
            # Converter para array se necessário
            if ($historico -isnot [Array]) { $historico = @($historico) }
        }
    }
    
    # Verificar se já existe snapshot do dia - se sim, atualizar
    $found = $false
    $newHistorico = @()
    foreach ($item in $historico) {
        if ($item.data -eq $hoje) {
            $newHistorico += $snapshot
            $found = $true
        } else {
            $newHistorico += $item
        }
    }
    if (-not $found) {
        $newHistorico += $snapshot
    }
    
    # Salvar
    $newHistorico | ConvertTo-Json -Depth 5 | Out-File -FilePath $historicoPath -Encoding UTF8
    
    # Salvar timestamp da última atualização
    $ultimaAtualizacaoPath = Join-Path $scriptDir "ultima-atualizacao.json"
    $ultimaAtualizacao = @{
        data = $hoje
        hora = (Get-Date -Format "HH:mm:ss")
        dataHoraCompleta = (Get-Date -Format "dd/MM/yyyy HH:mm:ss")
    }
    $ultimaAtualizacao | ConvertTo-Json | Out-File -FilePath $ultimaAtualizacaoPath -Encoding UTF8

    Write-Host "Snapshot gerado com sucesso!" -ForegroundColor Green
    Write-Host "  Data: $hoje $horaAtual" -ForegroundColor Cyan
    Write-Host "  Total cenarios: $total" -ForegroundColor Cyan
    foreach ($key in $statusCounts.Keys) {
        Write-Host "  $key : $($statusCounts[$key])" -ForegroundColor Cyan
    }
    
} catch {
    Write-Host "ERRO: $_" -ForegroundColor Red
    if ($excel) { $excel.Quit() }
    exit 1
} finally {
    if ($excel) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect()
    }
}
