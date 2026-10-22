# =========================================================
# PRE LOGETE
# V01
# =========================================================

Clear-Host

# =========================================================
# MAXIMIZAR JANELA
# =========================================================

try {

    Add-Type -AssemblyName System.Windows.Forms

    $process = Get-Process -Id $PID

    $sig = @'
        [DllImport("user32.dll")]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
'@

    $type = Add-Type `
        -MemberDefinition $sig `
        -Name Window `
        -Namespace Win32 `
        -PassThru

    $type::ShowWindowAsync($process.MainWindowHandle, 3)

}
catch {}

# =========================================================
# BUFFER
# =========================================================

try {

    $Host.UI.RawUI.BufferSize = `
        New-Object Management.Automation.Host.Size(180,5000)

}
catch {}

$margem = "    "

# =========================================================
# BANNER
# =========================================================

Write-Host ""
Write-Host ($margem + "██████╗ ██████╗ ███████╗") -ForegroundColor Cyan
Write-Host ($margem + "██╔══██╗██╔══██╗██╔════╝") -ForegroundColor Cyan
Write-Host ($margem + "██████╔╝██████╔╝█████╗") -ForegroundColor Cyan
Write-Host ($margem + "██╔═══╝ ██╔══██╗██╔══╝") -ForegroundColor Cyan
Write-Host ($margem + "██║     ██║  ██║███████╗") -ForegroundColor Cyan
Write-Host ($margem + "╚═╝     ╚═╝  ╚═╝╚══════╝") -ForegroundColor Cyan

Write-Host ""
Write-Host ($margem + "RELATORIO PRE LOGETE") -ForegroundColor Yellow
Write-Host ""

# =========================================================
# PASTA PROGRAMA
# =========================================================

if ($PSScriptRoot -and $PSScriptRoot.Trim() -ne "") {

    $pastaAtual = $PSScriptRoot
}
else {

    $pastaAtual = Split-Path `
        -Parent `
        ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
}

# =========================================================
# ARQUIVOS
# =========================================================

$caminhoCSV = `
"H:\DPR-DSU\PROGRAMAS\Pre Logete\bd_apontamentos_modelo_CSV_2026.csv"

$caminhoLista = Join-Path `
    $pastaAtual `
    "lista_recurso.txt"

# =========================================================
# VALIDACOES
# =========================================================

if (!(Test-Path $caminhoCSV)) {

    Write-Host ""
    Write-Host ($margem + "CSV nao encontrado!") -ForegroundColor Red
    Pause
    exit
}

if (!(Test-Path $caminhoLista)) {

    Write-Host ""
    Write-Host ($margem + "lista_recurso.txt nao encontrado!") -ForegroundColor Red
    Pause
    exit
}

# =========================================================
# IMPORTAR RECURSOS
# =========================================================

try {

    $textoLista = Get-Content `
        $caminhoLista `
        -Raw

    $recursosValidos = $textoLista.Split(";") |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -ne "" }

}
catch {

    Write-Host ""
    Write-Host ($margem + "Erro ao ler lista_recurso.txt") -ForegroundColor Red
    Pause
    exit
}

# =========================================================
# IMPORTAR CSV
# =========================================================

try {

    Write-Host ""
    Write-Host ($margem + "Carregando CSV...") -ForegroundColor Yellow

    $dados = Import-Csv $caminhoCSV

    Write-Host ($margem + "CSV carregado.") -ForegroundColor Green

}
catch {

    Write-Host ""
    Write-Host ($margem + "Erro ao importar CSV.") -ForegroundColor Red
    Write-Host ($margem + $_.Exception.Message)
    Pause
    exit
}

# =========================================================
# PROCESSAMENTO
# =========================================================

Write-Host ""
Write-Host ($margem + "Processando apontamentos...") -ForegroundColor Yellow

$ofsProcessadas = @{}
$resultados = @()

for ($i = $dados.Count - 1; $i -ge 0; $i--) {

    $linha = $dados[$i]

    $ofpc = $linha.'OF+PC'

    if ($ofsProcessadas.ContainsKey($ofpc)) {
        continue
    }

    $ofsProcessadas[$ofpc] = $true

    $recurso = $linha.RECURSO.Trim()

# =====================================================
# VALIDAR DATA (ULTIMOS 50 DIAS)
# =====================================================

try {

    $dataApontamento = [datetime]::Parse($linha.'DATA APONTAMENTO.1')

}
catch {

    continue
}

$limiteData = (Get-Date).AddDays(-50)

if ($dataApontamento -lt $limiteData) {

    continue
}

# =====================================================
# IGNORAR RETRABALHO
# =====================================================

$retrab = "$($linha.RETRABALHO)".Trim()

if ($retrab -ne "") {

    continue
}

    if ($recursosValidos -contains $recurso) {

        $resultados += [PSCustomObject]@{

            "OF+PC" = $ofpc
            "DATA"  = $linha.'DATA APONTAMENTO.1'
            "MODELO" = $linha.MODELO
        }
    }
}

# =========================================================
# ORDENAR
# =========================================================

$resultados = $resultados |
    Sort-Object 'OF+PC'

# =========================================================
# EXIBIR
# =========================================================

Write-Host ""
Write-Host ($margem + "TOTAL ENCONTRADO: $($resultados.Count)") -ForegroundColor Green
Write-Host ""

Write-Host ($margem + `
    ("{0,-15} {1,-20} {2,-20}" -f `
    "OF+PC",
    "DATA",
    "MODELO"))

Write-Host ($margem + ("-" * 60))

foreach ($item in $resultados) {

    Write-Host ($margem + `
        ("{0,-15} {1,-20} {2,-20}" -f `
        $item.'OF+PC',
        $item.DATA,
        $item.MODELO))
}

# =========================================================
# EXPORTAR
# =========================================================

Write-Host ""

$exportar = Read-Host ($margem + "Deseja gerar CSV? (S/N)")

if ($exportar.ToUpper() -eq "S") {

    try {

        $arquivoSaida = Join-Path `
            $pastaAtual `
            "pre_logete.csv"

        $resultados |
            Export-Csv `
                -Path $arquivoSaida `
                -NoTypeInformation `
                -Encoding UTF8

        Write-Host ""
        Write-Host ($margem + "CSV gerado com sucesso!") -ForegroundColor Green
        Write-Host ($margem + $arquivoSaida) -ForegroundColor Cyan

    }
    catch {

        Write-Host ""
        Write-Host ($margem + "Erro ao exportar CSV!") -ForegroundColor Red
        Write-Host ($margem + $_.Exception.Message)
    }
}

Write-Host ""
Write-Host ($margem + "Programa finalizado.") -ForegroundColor Yellow

Pause