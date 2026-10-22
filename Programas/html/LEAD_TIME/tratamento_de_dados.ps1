# ============================================================
# TRATAMENTO DE LEAD TIME
# Lê bd_lead_time.csv e gera trat_lead_time.csv
# Separador: ;
# Regra de seleção por ordem das linhas:
#   INÍCIO = primeira ocorrência de cima para baixo
#   FIM    = primeira ocorrência de baixo para cima
# ============================================================

$margem = "    "

Clear-Host

Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ($margem + "██╗     ███████╗ █████╗ ██████╗       ████████╗██╗███╗   ███╗███████╗") -ForegroundColor Cyan
Write-Host ($margem + "██║     ██╔════╝██╔══██╗██╔══██╗      ╚══██╔══╝██║████╗ ████║██╔════╝") -ForegroundColor Cyan
Write-Host ($margem + "██║     █████╗  ███████║██║  ██║█████╗   ██║   ██║██╔████╔██║█████╗  ") -ForegroundColor Cyan
Write-Host ($margem + "██║     ██╔══╝  ██╔══██║██║  ██║╚════╝   ██║   ██║██║╚██╔╝██║██╔══╝  ") -ForegroundColor Cyan
Write-Host ($margem + "███████╗███████╗██║  ██║██████╔╝         ██║   ██║██║ ╚═╝ ██║███████╗") -ForegroundColor Cyan
Write-Host ($margem + "╚══════╝╚══════╝╚═╝  ╚═╝╚═════╝          ╚═╝   ╚═╝╚═╝     ╚═╝╚══════╝") -ForegroundColor Cyan

Write-Host ""
Write-Host ($margem + "TRATAMENTO DE DADOS") -ForegroundColor Yellow
Write-Host ($margem + "Processamento automatico do Lead Time de producao") -ForegroundColor DarkGray
Write-Host ""
Write-Host ($margem + "Processando dados...") -ForegroundColor Green
Write-Host ""

# A partir daqui começa o seu programa normal
# -------------------------
# 1) CONFIGURAÇÕES GERAIS
# -------------------------

# Se o .ps1 estiver na mesma pasta do CSV, esta será a pasta usada.
# A pasta abaixo serve como fallback caso o script seja executado de outra forma.
$PastaPadrao = 'G:\DPR\AT\ALTA TENSÃO\SUPERVISOR\ESTAGIÁRIOS\09_ESTAGIARIOS_ATUAIS\Lucas Adhayr\1 - Projetos\Lead Time\excel'

if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $Pasta = $PSScriptRoot
}
else {
    $Pasta = $PastaPadrao
}

# Altere facilmente os nomes dos arquivos aqui
$NomeArquivoEntrada = 'bd_lead_time.csv'
$NomeArquivoSaida   = 'trat_lead_time.csv'

$ArquivoEntrada = Join-Path $Pasta $NomeArquivoEntrada
$ArquivoSaida   = Join-Path $Pasta $NomeArquivoSaida

$Separador = ';'
$FormatoDataSaida = 'd/M/yy H:mm'

# -------------------------
# 2) FASES E RECURSOS
# -------------------------
# Para adicionar uma nova fase, basta copiar um bloco e trocar o nome/recursos.
# Os recursos devem ficar entre aspas.

$Fases = [ordered]@{
    'BS/BP' = @{
        Inicio = @('101','102','103','104','105','106','107','301','302','303','304','305','306','307','1101')
        Fim    = @('701','702','703','704','1101')
    }

    'IS/RB' = @{
        Inicio = @('801','802','1201')
        Fim    = @('801','802','901','1001','1501')
    }

    'FD/MN' = @{
        Inicio = @('1820','1601')
        Fim    = @('1870','1720','1721','1722','1723','1724')
    }

    'MP' = @{
        Inicio = @('1900','1910','2001')
        Fim    = @('2101','2102','2103','2104','2105','2001')
    }

    'TO' = @{
        Inicio = @('3001','3010')
        Fim    = @('3100')
    }

    'AC' = @{
        Inicio = @('3200')
        Fim    = @('3267','3241','3242','3243','3244','3245')
    }

    'AF' = @{
        Inicio = @('3301','3302','3303')
        Fim    = @('3500','3400')
    }
}

# -------------------------
# 3) EXCEÇÃO DE MODELOS
# -------------------------
# Estes modelos não possuem a fase FD/MN.
# Para eles, INÍCIO FD/MN e FIM FD/MN receberão os mesmos valores de IS/RB.

$ModelosSemFDMN = @(
    'TCM-145',
    'TCM-145Q',
    'TCM-145T',
    'TCM-245',
    'TCM-362',
    'TCM-362S',
    'TCM-420',
    'TCM-550',
    'TCM-72',
    'TCM-72S'
)

# ============================================================
# FUNÇÕES
# ============================================================

function Converter-Data {
    param(
        [AllowNull()]
        [string]$Texto
    )

    if ([string]::IsNullOrWhiteSpace($Texto)) {
        return $null
    }

    $Texto = $Texto.Trim()
    $Cultura = [System.Globalization.CultureInfo]::GetCultureInfo('pt-BR')
    $Estilo = [System.Globalization.DateTimeStyles]::AllowWhiteSpaces

    # Formatos esperados no CSV + algumas variações comuns
    $Formatos = @(
        'dd/MM/yyyy HH:mm',
        'd/MM/yyyy HH:mm',
        'dd/M/yyyy HH:mm',
        'd/M/yyyy HH:mm',
        'dd/MM/yyyy H:mm',
        'd/MM/yyyy H:mm',
        'dd/M/yyyy H:mm',
        'd/M/yyyy H:mm',
        'dd/MM/yyyy',
        'd/MM/yyyy',
        'dd/M/yyyy',
        'd/M/yyyy'
    )

    foreach ($Formato in $Formatos) {
        $Data = [datetime]::MinValue
        if ([datetime]::TryParseExact($Texto, $Formato, $Cultura, $Estilo, [ref]$Data)) {
            return $Data
        }
    }

    # Última tentativa, usando o padrão pt-BR
    $Data = [datetime]::MinValue
    if ([datetime]::TryParse($Texto, $Cultura, $Estilo, [ref]$Data)) {
        return $Data
    }

    Write-Warning "Data inválida ignorada: '$Texto'"
    return $null
}

function Formatar-Data {
    param(
        [AllowNull()]
        $Data
    )

    if ($null -eq $Data) {
        return ''
    }

    return ([datetime]$Data).ToString($FormatoDataSaida, [System.Globalization.CultureInfo]::InvariantCulture)
}

# ============================================================
# PREPARAÇÃO
# ============================================================

if (-not (Test-Path -LiteralPath $ArquivoEntrada)) {
    Write-Host "ERRO: arquivo de entrada não encontrado:" -ForegroundColor Red
    Write-Host $ArquivoEntrada -ForegroundColor Red
    exit 1
}

# Lookup: recurso -> fase(s) em que ele representa INÍCIO
$LookupInicio = @{}

# Lookup: recurso -> fase(s) em que ele representa FIM
$LookupFim = @{}

foreach ($NomeFase in $Fases.Keys) {
    foreach ($Recurso in $Fases[$NomeFase].Inicio) {
        $ChaveRecurso = ([string]$Recurso).Trim()

        if (-not $LookupInicio.ContainsKey($ChaveRecurso)) {
            $LookupInicio[$ChaveRecurso] = @()
        }

        $LookupInicio[$ChaveRecurso] += $NomeFase
    }

    foreach ($Recurso in $Fases[$NomeFase].Fim) {
        $ChaveRecurso = ([string]$Recurso).Trim()

        if (-not $LookupFim.ContainsKey($ChaveRecurso)) {
            $LookupFim[$ChaveRecurso] = @()
        }

        $LookupFim[$ChaveRecurso] += $NomeFase
    }
}

# OrderedDictionary mantém a ordem da primeira aparição de cada OF_PC no arquivo.
$DadosPorOFPC = New-Object System.Collections.Specialized.OrderedDictionary

$PrimeiraLinha = $true
$LinhasProcessadas = 0

# ============================================================
# LEITURA E AGREGAÇÃO DO CSV
# ============================================================

try {
    Import-Csv -LiteralPath $ArquivoEntrada -Delimiter $Separador -Encoding Default | ForEach-Object {
        $Linha = $_
        $LinhasProcessadas++

        if ($PrimeiraLinha) {
            $ColunasObrigatorias = @(
                'OF_PC',
                'MODELO.1',
                'DT/HORA INICIO',
                'DT/HR FIM.1',
                'Recurso'
            )

            $NomesColunas = @($Linha.PSObject.Properties.Name)

            foreach ($Coluna in $ColunasObrigatorias) {
                if ($NomesColunas -notcontains $Coluna) {
                    throw "A coluna obrigatória '$Coluna' não foi encontrada no CSV."
                }
            }

            $PrimeiraLinha = $false
        }

        $OFPC = ([string]$Linha.OF_PC).Trim()

        if ([string]::IsNullOrWhiteSpace($OFPC)) {
            return
        }

        $Modelo = ([string]$Linha.'MODELO.1').Trim()

        if (-not $DadosPorOFPC.Contains($OFPC)) {
            $DadosFases = @{}

            foreach ($NomeFase in $Fases.Keys) {
                $DadosFases[$NomeFase] = [pscustomobject]@{
                    Inicio = $null
                    Fim    = $null
                }
            }

            $DadosPorOFPC.Add(
                $OFPC,
                [pscustomobject]@{
                    Modelo = $Modelo
                    Fases  = $DadosFases
                }
            )
        }

        $RegistroOF = $DadosPorOFPC[$OFPC]

        # Se o primeiro registro daquele OF_PC estiver sem modelo, tenta preencher depois.
        if ([string]::IsNullOrWhiteSpace($RegistroOF.Modelo) -and -not [string]::IsNullOrWhiteSpace($Modelo)) {
            $RegistroOF.Modelo = $Modelo
        }

        $RecursoAtual = ([string]$Linha.Recurso).Trim()

        # ---------- INÍCIO DE FASE ----------
        # O CSV está em ordem cronológica (mais antigo -> mais novo).
        # Para o INÍCIO, queremos a PRIMEIRA ocorrência encontrada de cima para baixo.
        # Portanto, depois que uma fase recebe um início, ele NÃO é substituído.
        if ($LookupInicio.ContainsKey($RecursoAtual)) {
            $DataInicio = Converter-Data $Linha.'DT/HORA INICIO'

            if ($null -ne $DataInicio) {
                foreach ($NomeFase in $LookupInicio[$RecursoAtual]) {
                    if ($null -eq $RegistroOF.Fases[$NomeFase].Inicio) {
                        $RegistroOF.Fases[$NomeFase].Inicio = $DataInicio
                    }
                }
            }
        }

        # ---------- FIM DE FASE ----------
        # Para o FIM, queremos a PRIMEIRA ocorrência encontrada de baixo para cima.
        # Como o arquivo está sendo lido de cima para baixo, basta SOBRESCREVER o fim
        # sempre que aparecer um novo recurso de fim. Ao terminar a leitura, ficará
        # armazenada a ocorrência mais baixa do arquivo (a última de cima para baixo).
        if ($LookupFim.ContainsKey($RecursoAtual)) {
            $DataFim = Converter-Data $Linha.'DT/HR FIM.1'

            if ($null -ne $DataFim) {
                foreach ($NomeFase in $LookupFim[$RecursoAtual]) {
                    $RegistroOF.Fases[$NomeFase].Fim = $DataFim
                }
            }
        }

        if (($LinhasProcessadas % 10000) -eq 0) {
            Write-Progress -Activity 'Processando bd_lead_time.csv' -Status "$LinhasProcessadas linhas processadas..."
        }
    }
}
catch {
    Write-Progress -Activity 'Processando bd_lead_time.csv' -Completed
    Write-Host "ERRO ao processar o CSV:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

Write-Progress -Activity 'Processando bd_lead_time.csv' -Completed

if ($PrimeiraLinha) {
    Write-Host 'ERRO: o CSV está vazio ou não possui registros de dados.' -ForegroundColor Red
    exit 1
}

# ============================================================
# REGRA ESPECIAL: MODELOS SEM FD/MN
# ============================================================

if (($Fases.Keys -contains 'IS/RB') -and ($Fases.Keys -contains 'FD/MN')) {
    foreach ($OFPC in $DadosPorOFPC.Keys) {
        $RegistroOF = $DadosPorOFPC[$OFPC]

        if ($ModelosSemFDMN -contains $RegistroOF.Modelo) {
            $RegistroOF.Fases['FD/MN'].Inicio = $RegistroOF.Fases['IS/RB'].Inicio
            $RegistroOF.Fases['FD/MN'].Fim    = $RegistroOF.Fases['IS/RB'].Fim
        }
    }
}

# ============================================================
# MONTAGEM E EXPORTAÇÃO DO RESULTADO
# ============================================================

$Resultado = foreach ($OFPC in $DadosPorOFPC.Keys) {
    $RegistroOF = $DadosPorOFPC[$OFPC]

    $LinhaSaida = [ordered]@{
        'OF_PC'  = $OFPC
        'MODELO' = $RegistroOF.Modelo
    }

    foreach ($NomeFase in $Fases.Keys) {
        $LinhaSaida["INÍCIO $NomeFase"] = Formatar-Data $RegistroOF.Fases[$NomeFase].Inicio
        $LinhaSaida["FIM $NomeFase"]    = Formatar-Data $RegistroOF.Fases[$NomeFase].Fim
    }

    [pscustomobject]$LinhaSaida
}

try {
    # Export-Csv sem -Append já sobrescreve o arquivo existente.
    # -Force também permite substituir o arquivo quando possível.
    $Resultado | Export-Csv -LiteralPath $ArquivoSaida -Delimiter $Separador -NoTypeInformation -Encoding UTF8 -Force
}
catch {
    Write-Host "ERRO ao gerar o arquivo de saída:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host 'Verifique se trat_lead_time.csv não está aberto no Excel.' -ForegroundColor Yellow
    exit 1
}

Write-Host ''
Write-Host 'Processamento concluído com sucesso.' -ForegroundColor Green
Write-Host "Linhas lidas: $LinhasProcessadas"
Write-Host "OF_PC únicos: $($DadosPorOFPC.Count)"
Write-Host "Arquivo gerado: $ArquivoSaida"
