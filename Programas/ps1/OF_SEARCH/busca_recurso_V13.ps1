# =========================================================
# CONSULTA DE APONTAMENTOS POR OF+PC
# V05.1 - VISUAL MELHORADO
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

    # 3 = Maximizado
    $type::ShowWindowAsync($process.MainWindowHandle, 3)

}
catch {}

# =========================================================
# TAMANHO BUFFER
# =========================================================

try {

    $Host.UI.RawUI.BufferSize = `
        New-Object Management.Automation.Host.Size(220,5000)

}
catch {}

# =========================================================
# MARGEM ESQUERDA
# =========================================================

$margem = "    "

# =========================================================
# BANNER
# =========================================================
Write-Host ""
Write-Host ($margem + "█    █████╗ ████████╗         ███╗   ███╗████████╗  █") -ForegroundColor DarkYellow
Write-Host ($margem + "█   ██╔══██╗╚══██╔══╝         ████╗ ████║╚══██╔══╝  █") -ForegroundColor DarkYellow
Write-Host ($margem + "█   ███████║   ██║    ████╗   ██╔████╔██║   ██║     █") -ForegroundColor DarkYellow
Write-Host ($margem + "█   ██╔══██║   ██║    ╚═══╝   ██║╚██╔╝██║   ██║     █") -ForegroundColor DarkYellow
Write-Host ($margem + "█   ██║  ██║   ██║            ██║ ╚═╝ ██║   ██║     █") -ForegroundColor DarkYellow
Write-Host ($margem + "█   ╚═╝  ╚═╝   ╚═╝            ╚═╝     ╚═╝   ╚═╝     █") -ForegroundColor DarkYellow
Write-Host ""
Write-Host ($margem + "█  ██╗    ██╗ ███████╗  ██████╗     █") -ForegroundColor darkBlue
Write-Host ($margem + "█  ██║    ██║ ██╔════╝ ██╔════╝     █") -ForegroundColor darkBlue
Write-Host ($margem + "█  ██║ █╗ ██║ █████╗   ██║  ███╗    █") -ForegroundColor darkBlue
Write-Host ($margem + "█  ██║███╗██║ ██╔══╝   ██║   ██║    █   BALTEAU PRODUTOS ELÉTRICOS LTDA") -ForegroundColor darkBlue
Write-Host ($margem + "█  ╚███╔███╔╝ ███████╗ ╚██████╔╝    █   SISTEMA DE VISUALIZAÇÃO DE APONTAMENTOS") -ForegroundColor darkBlue
Write-Host ($margem + "█   ╚══╝╚══╝  ╚══════╝  ╚═════╝     █   OF_SEARCH") -ForegroundColor darkBlue



# =========================================================
# CAMINHO DA PASTA
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

$caminhoCSV = Join-Path `
    $pastaAtual `
    "bd_apontamentos_modelo_CSV_2026.csv"

$caminhoFluxo = Join-Path `
    $pastaAtual `
    "modelo_recurso.txt"

# =========================================================
# VALIDAR CSV
# =========================================================

if (!(Test-Path $caminhoCSV)) {

    Write-Host ""
    Write-Host ($margem + "ERRO: CSV nao encontrado!") -ForegroundColor DarkRed
    Write-Host ($margem + $caminhoCSV)
    Pause
    exit
}

# =========================================================
# VALIDAR FLUXO
# =========================================================

if (!(Test-Path $caminhoFluxo)) {

    Write-Host ""
    Write-Host ($margem + "ERRO: modelo_recurso.txt nao encontrado!") -ForegroundColor DarkRed
    Write-Host ($margem + $caminhoFluxo)
    Pause
    exit
}

# =========================================================
# IMPORTAR CSV
# =========================================================

try {

    Write-Host ""
    Write-Host ($margem + "Carregando CSV...") -ForegroundColor Cyan

    $dados = Import-Csv $caminhoCSV

    $dados = $dados | Where-Object {
        $_.'OF+PC' -ne $null
    }

    Write-Host ($margem + "CSV carregado com sucesso!") -ForegroundColor Green
}
catch {

    Write-Host ""
    Write-Host ($margem + "ERRO AO IMPORTAR CSV!") -ForegroundColor Red
    Write-Host ($margem + $_.Exception.Message)
    Pause
    exit
}

# =========================================================
# DICIONARIO DESCRICOES
# =========================================================

$descricaoRecursos = @{}

foreach ($linha in $dados) {

    $r = $linha.RECURSO.Trim()

    if (!$descricaoRecursos.ContainsKey($r)) {

        $descricaoRecursos[$r] = $linha.DESCRICAO.Trim()
    }
}

# =========================================================
# IMPORTAR FLUXOS
# =========================================================

$fluxos = @{}

$linhasFluxo = Get-Content $caminhoFluxo

$modelosAtuais = @()

foreach ($linha in $linhasFluxo) {

    $linha = $linha.Trim()

    if ($linha -eq "") {
        continue
    }

    # =====================================================
    # MODELOS
    # =====================================================

    if ($linha -match "^[A-Z]") {

        $modelosAtuais = $linha -split ";"

        foreach ($m in $modelosAtuais) {

            $modelo = $m.Trim()

            if ($modelo -ne "") {

                $fluxos[$modelo] = @()
            }
        }

        continue
    }

    # =====================================================
    # OPCIONAL
    # =====================================================

    $opcional = $false

    if ($linha.StartsWith("{")) {

        $opcional = $true

        $linha = $linha.Replace("{", "")
        $linha = $linha.Replace("}", "")
    }

    $recursos = $linha -split ";" | Where-Object {
        $_.Trim() -ne ""
    }

    $recursos = $recursos | Sort-Object {
        [int]$_
    }

    $etapa = [PSCustomObject]@{
        Recursos = $recursos
        Opcional = $opcional
    }

    foreach ($modelo in $modelosAtuais) {

        $fluxos[$modelo] += $etapa
    }
}

# =========================================================
# EXPORTACAO
# =========================================================

$resultadosExportacao = @()

# =========================================================
# LOOP PRINCIPAL
# =========================================================

do {

    Write-Host ""
    Write-Host ($margem + "Digite a OF-PC para pesquisar uma peça específica") -ForegroundColor Cyan
    Write-Host ($margem + "Ou digite somente a OF para pesquisar a OF completa") -ForegroundColor DarkCyan
    Write-Host ""
    Write-Host ($margem + "Exemplos:") -ForegroundColor DarkGray
    Write-Host ($margem + "255134-2") -ForegroundColor DarkGray
    Write-Host ($margem + "255134-2;255537-31;265166-18...") -ForegroundColor DarkGray
    Write-Host ($margem + "255444") -ForegroundColor DarkGray
    Write-Host ($margem + "255444;255134;265001...") -ForegroundColor DarkGray
    Write-Host ""

    $entrada = Read-Host ($margem + "PESQUISA") 

    $consultas = $entrada -split ";"

    foreach ($consulta in $consultas) {

        $consulta = $consulta.Trim()

        if ($consulta -eq "") {
            continue
        }

        # =====================================================
        # OF+PC
        # =====================================================

        if ($consulta -like "*-*") {

            $ofsEncontradas = @($consulta)
        }

        # =====================================================
        # SOMENTE OF
        # =====================================================

        else {

            $ofsEncontradas = $dados |
                Where-Object {
                    $_.OF.Trim() -eq $consulta
                } |
                Select-Object -ExpandProperty 'OF+PC' -Unique |
                Sort-Object
        }

        # =====================================================
        # NADA
        # =====================================================

        if ($ofsEncontradas.Count -eq 0) {

            Write-Host ""
            Write-Host ($margem + "$consulta -> OF NÃO ENCONTRADA") -ForegroundColor yellow
            continue
        }

        foreach ($of in $ofsEncontradas) {

            
            Write-Host ($margem + "------------------------------------------------------------") -ForegroundColor DarkGray
            Write-Host ""

            $resultado = $dados | Where-Object {
                $_.'OF+PC'.Trim() -eq $of
            }

            if ($resultado.Count -eq 0) {
                continue
            }

            $modelo = $resultado[0].MODELO.Trim()

            if (!$fluxos.ContainsKey($modelo)) {

                Write-Host ""
                Write-Host ($margem + "O fluxo para o modelo $modelo não foi encontrado") -ForegroundColor Red
                Write-Host "Verifique no documento modelo_recurso se ele existe ou adicione"
                continue
            }

            # =================================================
            # TITULO CENTRALIZADO
            # =================================================

            $titulo = "$of [$modelo]"

            $largura = 120

            $espacos = [Math]::Floor(($largura - $titulo.Length) / 2)

            Write-Host (($margem) + (" " * $espacos) + $titulo) -ForegroundColor Green

            Write-Host ""

            # =================================================
            # APONTAMENTOS
            # =================================================

            $apontamentos = @{}
            $repetidos = @{}

            foreach ($linha in $resultado) {

                $r = $linha.RECURSO.Trim()

                if (!$apontamentos.ContainsKey($r)) {

                    $apontamentos[$r] = $linha
                }
                else {

                    if (!$repetidos.ContainsKey($r)) {
                        $repetidos[$r] = @()
                    }

                    $repetidos[$r] += $linha
                }
            }

            $ultimoApontamento = $resultado[-1].RECURSO.Trim()

            # =================================================
            # CABECALHO
            # =================================================

            Write-Host ($margem + `
                ("{0,-12} {1,-45} {2,-20} {3,-12} {4,-10} {5,-35}" -f `
                "RECURSO",
                "DESCRICAO",
                "DATA",
                "OPERADOR",
                "STATUS",
                "OBSERVACAO"))

            Write-Host ($margem + ("-" * 145))

            Write-Host ""

            # =================================================
            # CONTROLE FLUXO
            # V07 - CORRECAO: etapa nao apontada nao contamina
            # o proximo recurso com FLUXO INCORRETO.
            # FLUXO INCORRETO so ocorre quando um recurso foi
            # apontado em data anterior a um recurso obrigatorio
            # anterior que tambem foi apontado.
            # =================================================

            # Monta lista ordenada de recursos obrigatorios
            


$pararFluxo = $false


            foreach ($etapa in $fluxos[$modelo]) {

                if ($pararFluxo) {
                    break
                }

                # =================================================
                # OPCIONAL
                # =================================================

                if ($etapa.Opcional) {

                    foreach ($r in $etapa.Recursos) {

                        if ($apontamentos.ContainsKey($r)) {

                            $linha = $apontamentos[$r]

                            $status = "OK"
                            $obs = ""

                            if ($r -eq $ultimoApontamento) {

                                $obs = "ULTIMO APONTAMENTO"
                                $pararFluxo = $true
                            }

                            $texto = $margem + `
                                ("{0,-12} {1,-45} {2,-20} {3,-12} {4,-10} {5,-35}" -f `
                                $r,
                                $linha.DESCRICAO,
                                $linha.'DATA APONTAMENTO.1',
                                $linha.OPERADOR,
                                $status,
                                $obs)

                            if ($obs -like "*ULTIMO*") {

                                Write-Host $texto -ForegroundColor Green
                            }
                            else {

                                Write-Host $texto
                            }
                        }
                    }

                    continue
                }

                # =================================================
                # OBRIGATORIOS
                # =================================================

                foreach ($r in $etapa.Recursos) {

                    $status = "OK"
                    $obs = ""

                    $descricao = "SEM DESCRICAO"

                    if ($descricaoRecursos.ContainsKey($r)) {

                        $descricao = $descricaoRecursos[$r]
                    }

                    # =============================================
                    # NAO APONTADO
                    # =============================================

                    if (!$apontamentos.ContainsKey($r)) {

    $status = "ERRO"
    $obs = "NAO APONTADO"

    $texto = $margem + `
        ("{0,-12} {1,-45} {2,-20} {3,-12} {4,-10} {5,-35}" -f `
        $r,
        $descricao,
        "SEM DATA",
        "SEM OPERADOR",
        $status,
        $obs)

    Write-Host $texto -ForegroundColor DarkRed


    # =============================================
    # EXPORTAR NAO APONTADO
    # =============================================

    $resultadosExportacao += [PSCustomObject]@{

        "OF+PC"      = $of
        "MODELO"     = $modelo
        "RECURSO"    = $r
        "DESCRICAO"  = $descricao
        "DATA"       = "SEM DATA"
        "OPERADOR"   = "SEM OPERADOR"
        "STATUS"     = $status
        "OBSERVACAO" = $obs
    }


    continue
}

                    # =============================================
                    # APONTADO - VERIFICA ORDEM CRONOLOGICA
                    # =============================================

$linha = $apontamentos[$r]

$obs = ""

$status = "OK"

                    # =============================================
                    # REPETIDO
                    # =============================================

                    if ($repetidos.ContainsKey($r)) {

                        if ($obs -ne "") {
                            $obs += " | "
                        }

                        $obs += "REPETIDO"

                        $status = "ERRO"
                    }

                    # =============================================
                    # ULTIMO
                    # =============================================

                    if ($r -eq $ultimoApontamento) {

                        if ($obs -ne "") {
                            $obs += " | "
                        }

                        $obs += "ULTIMO APONTAMENTO"

                        $pararFluxo = $true
                    }

                    $texto = $margem + `
                        ("{0,-12} {1,-45} {2,-20} {3,-12} {4,-10} {5,-35}" -f `
                        $r,
                        $linha.DESCRICAO,
                        $linha.'DATA APONTAMENTO.1',
                        $linha.OPERADOR,
                        $status,
                        $obs)

                    if ($status -eq "ERRO") {

                        Write-Host $texto -ForegroundColor Red
                    }
                    elseif ($obs -like "*ULTIMO*") {

                        Write-Host $texto -ForegroundColor Green
                    }
                    else {

                        Write-Host $texto
                    }

                    # =============================================
                    # EXPORTAR
                    # =============================================

                    $resultadosExportacao += [PSCustomObject]@{

                        "OF+PC"      = $of
                        "MODELO"     = $modelo
                        "RECURSO"    = $r
                        "DESCRICAO"  = $linha.DESCRICAO
                        "DATA"       = $linha.'DATA APONTAMENTO.1'
                        "OPERADOR"   = $linha.OPERADOR
                        "STATUS"     = $status
                        "OBSERVACAO" = $obs
                    }
                }
            }
        }
    }

    Write-Host ""

    $continuar = Read-Host ($margem + "Deseja pesquisar mais OF+PC? (S/N)")

} while ($continuar.ToUpper() -eq "S")

# =========================================================
# EXPORTAR CSV
# =========================================================

Write-Host ""

$exportar = Read-Host ($margem + "Deseja gerar CSV? (S/N)")

if ($exportar.ToUpper() -eq "S") {

    try {

        $dataHora = Get-Date -Format "yyyyMMdd_HHmmss"

        # =====================================================
        # PASTA FIXA DOS RELATORIOS
        # =====================================================

        $pastaRelatorios = `
        "H:\DPR-DSU\PROGRAMAS\OF_Search\RELATORIOS"

        # =====================================================
        # CRIA PASTA SE NAO EXISTIR
        # =====================================================

        if (!(Test-Path $pastaRelatorios)) {

            New-Item `
                -ItemType Directory `
                -Path $pastaRelatorios | Out-Null
        }

        # =====================================================
        # NOME ARQUIVO
        # =====================================================

        $arquivoSaida = Join-Path `
            $pastaRelatorios `
            "resultado_fluxo_$dataHora.csv"

        # =====================================================
        # EXPORTAR
        # =====================================================

        $resultadosExportacao |
            Export-Csv `
                -Path $arquivoSaida `
                -NoTypeInformation `
                -Encoding UTF8

        Write-Host ""
        Write-Host ($margem + "CSV exportado com sucesso!") -ForegroundColor Green
        Write-Host ($margem + $arquivoSaida) -ForegroundColor Cyan
    }
    catch {

        Write-Host ""
        Write-Host ($margem + "ERRO ao exportar CSV!") -ForegroundColor Red
        Write-Host ($margem + $_.Exception.Message)
    }
}

Write-Host ""
Write-Host ($margem + "Programa finalizado.") -ForegroundColor Yellow

Pause