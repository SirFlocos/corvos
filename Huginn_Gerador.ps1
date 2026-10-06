# ============================================================
#  Huginn_Gerador.ps1 - Gerador de Massa de Testes para o MUNIN
#  Uso: duplo-clique no Hugin.bat
# ============================================================

$ErrorActionPreference = 'Stop'
$Root      = Split-Path -Parent $MyInvocation.MyCommand.Path
$PastaXml  = Join-Path $Root '01_XML_Fornecedores'
$PastaSped = Join-Path $Root '02_Dados_ERP'
$Marker    = Join-Path $Root '.huginn_demo'

if (-not (Test-Path $PastaXml))  { New-Item -ItemType Directory -Path $PastaXml  | Out-Null }
if (-not (Test-Path $PastaSped)) { New-Item -ItemType Directory -Path $PastaSped | Out-Null }

$CNPJ_PADRAO = '11222333000181'

function Test-CnpjValido {
    param([string]$Valor)
    if ([string]::IsNullOrWhiteSpace($Valor)) { return $null }
    $s = ($Valor -replace '\W','').ToUpper()
    if ($s.Length -ne 14) { return $null }
    $b  = $s.Substring(0,12)
    $dv = $s.Substring(12,2)
    if (($b -replace '[0-9A-Z]','') -ne '') { return $null }
    if (($dv -replace '[0-9]','') -ne '')   { return $null }
    $p = 5,4,3,2,9,8,7,6,5,4,3,2
    $t = 0
    for ($i = 0; $i -lt 12; $i++) { $t += (([int][char]$b[$i]) - 48) * $p[$i] }
    $r  = $t % 11
    if ($r -lt 2) { $d1 = 0 } else { $d1 = 11 - $r }
    $p = 6,5,4,3,2,9,8,7,6,5,4,3,2
    $t = 0
    for ($i = 0; $i -lt 12; $i++) { $t += (([int][char]$b[$i]) - 48) * $p[$i] }
    $t += $d1 * 2
    $r  = $t % 11
    if ($r -lt 2) { $d2 = 0 } else { $d2 = 11 - $r }
    if (([string]$d1 + [string]$d2) -eq $dv) { return $s }
    return $null
}

Clear-Host
Write-Host ""
Write-Host " ============================================"
Write-Host "   H.U.G.I.N.N."
Write-Host "   Gerador de Massa de Testes para o MUNIN"
Write-Host " ============================================"
Write-Host ""
Write-Host "  Este programa vai gerar XMLs e um arquivo SPED"
Write-Host "  de teste, com erros propositais, para validar"
Write-Host "  a auditoria do MUNIN."
Write-Host ""
Write-Host " --------------------------------------------"
Write-Host "  QUAL O CNPJ DA EMPRESA A SER AUDITADA?"
Write-Host " --------------------------------------------"
Write-Host ""
Write-Host "  - Digite os 14 numeros do CNPJ e aperte ENTER"
Write-Host "  - Ou apenas aperte ENTER para usar o padrao:"
Write-Host "    $CNPJ_PADRAO"
Write-Host ""

$cnpjAuditado = $null
while (-not $cnpjAuditado) {
    $entrada = Read-Host " CNPJ"

    if ([string]::IsNullOrWhiteSpace($entrada)) {
        $cnpjAuditado = $CNPJ_PADRAO
        Write-Host ""
        Write-Host "  >> Usando CNPJ padrao: $cnpjAuditado" -ForegroundColor Cyan
    }
    else {
        $cnpjAuditado = Test-CnpjValido $entrada
        if (-not $cnpjAuditado) {
            Write-Host ""
            Write-Host "  [X] Esse CNPJ nao e valido." -ForegroundColor Red
            Write-Host "      Confira se tem 14 digitos e se os" -ForegroundColor Red
            Write-Host "      digitos verificadores estao corretos." -ForegroundColor Red
            Write-Host ""
            Write-Host "      Tente de novo, ou aperte ENTER para"
            Write-Host "      usar o padrao $CNPJ_PADRAO"
            Write-Host ""
            $cnpjAuditado = $null
        }
    }
}

Write-Host ""
Write-Host " --------------------------------------------"
Write-Host "  Gerando massa de testes, aguarde..."
Write-Host " --------------------------------------------"
Write-Host ""

# ---------- Limpeza ----------
Get-ChildItem -Path $PastaXml  -Filter "*.xml" -ErrorAction SilentlyContinue | Remove-Item -Force
Get-ChildItem -Path $PastaSped -Filter "*.txt" -ErrorAction SilentlyContinue | Remove-Item -Force

$spedLines = [System.Collections.Generic.List[string]]::new()
$spedLines.Add("|0000|014|0|01092026|30092026|EMPRESA TESTE AUDITORIA LTDA|$cnpjAuditado||PB|2507507|||A|1|")

$totalNotas = Get-Random -Minimum 80 -Maximum 120
$indices    = 1..$totalNotas | Get-Random -Count $totalNotas

$qtdDivergenciaValor = Get-Random -Minimum 3 -Maximum 6
$qtdOmissaoSped      = Get-Random -Minimum 2 -Maximum 5
$qtdOmissaoXml       = Get-Random -Minimum 2 -Maximum 4
$qtdSerieDivergente  = Get-Random -Minimum 1 -Maximum 3

$idx = 0
$idxDivergValor = $indices[$idx..($idx + $qtdDivergenciaValor - 1)]; $idx += $qtdDivergenciaValor
$idxOmissaoSped = $indices[$idx..($idx + $qtdOmissaoSped - 1)];     $idx += $qtdOmissaoSped
$idxOmissaoXml  = $indices[$idx..($idx + $qtdOmissaoXml - 1)];      $idx += $qtdOmissaoXml
$idxSerieDiverg = $indices[$idx..($idx + $qtdSerieDivergente - 1)]; $idx += $qtdSerieDivergente

for ($i = 1; $i -le $totalNotas; $i++) {

    $iStr   = '{0:D3}' -f $i
    $iChave = '{0:D9}' -f $i
    $chave  = '352609' + $cnpjAuditado + '55001' + $iChave + '1012345678'

    $vItem     = [math]::Round((Get-Random -Minimum 100.0 -Maximum 3000.0), 2)
    $vSped     = $vItem
    $serieSped = '1'

    $escreverXml  = $true
    $escreverSped = $true

    if ($i -in $idxDivergValor)     { $vSped = [math]::Round($vItem - 50.00, 2) }
    elseif ($i -in $idxOmissaoSped) { $escreverSped = $false }
    elseif ($i -in $idxOmissaoXml)  { $escreverXml  = $false }
    elseif ($i -in $idxSerieDiverg) { $serieSped    = '2'    }

    $strVItemXml = $vItem.ToString('0.00', [System.Globalization.CultureInfo]::InvariantCulture)
    $strVSpedTxt = $vSped.ToString('0.00', [System.Globalization.CultureInfo]::InvariantCulture).Replace('.', ',')

    if ($escreverXml) {
        $xmlContent = '<?xml version="1.0" encoding="UTF-8"?>' +
            '<nfeProc xmlns="http://www.portalfiscal.inf.br/nfe" versao="4.00">' +
            '<NFe><infNFe Id="NFe' + $chave + '">' +
            '<ide><nNF>' + $i + '</nNF><serie>1</serie><tpNF>1</tpNF><dhEmi>2026-09-25T10:00:00-03:00</dhEmi></ide>' +
            '<emit><CNPJ>98765432000188</CNPJ><xNome>FORNECEDOR TESTE MUNIN ' + $iStr + ' LTDA</xNome></emit>' +
            '<dest><CNPJ>' + $cnpjAuditado + '</CNPJ></dest>' +
            '<det nItem="1"><prod><cProd>PROD' + $iStr + '</cProd><xProd>PRODUTO TESTE ' + $iStr + '</xProd><CFOP>5102</CFOP><uCom>UN</uCom><qCom>1.0000</qCom><vUnCom>' + $strVItemXml + '</vUnCom><vProd>' + $strVItemXml + '</vProd></prod></det>' +
            '<total><ICMSTot><vNF>' + $strVItemXml + '</vNF><vProd>' + $strVItemXml + '</vProd><vFrete>0.00</vFrete></ICMSTot></total>' +
            '</infNFe></NFe></nfeProc>'

        [System.IO.File]::WriteAllText((Join-Path $PastaXml ('NFe_' + $chave + '.xml')), $xmlContent, [System.Text.Encoding]::UTF8)
    }

    if ($escreverSped) {
        $c100 = "|C100|0|1|FORN$iStr|55|00|$serieSped|$i|$chave|25092026|25092026|$strVSpedTxt|1|0,00|0,00|$strVSpedTxt|0|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|"
        $c170 = "|C170|1|PROD$iStr|PRODUTO TESTE $iStr|1,0000|UN|$strVSpedTxt|0,00|0|0|5102||$strVSpedTxt|"
        $spedLines.Add($c100)
        $spedLines.Add($c170)
    }
}

$spedLines.Add('|9999|99|')
[System.IO.File]::WriteAllLines((Join-Path $PastaSped 'SPED_AUDITORIA_TESTE.txt'), $spedLines, [System.Text.Encoding]::GetEncoding('iso-8859-1'))

# ---------- Marker para o Ymir detectar modo demo ----------
# Formato chave=valor, UTF-8 sem BOM.
# O Ymir.bat le isso e pergunta se o usuario quer rodar em modo demo.
$markerLinhas = @(
    "CNPJ=$cnpjAuditado"
    "DATA_INICIO=01/09/2026"
    "DATA_FIM=30/09/2026"
    "GERADO_EM=$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
)
[System.IO.File]::WriteAllLines($Marker, $markerLinhas)

$totalXmlsGerados = $totalNotas - $qtdOmissaoXml
$totalSpedGerados = $totalNotas - $qtdOmissaoSped
$qtdConforme      = $totalNotas - ($qtdDivergenciaValor + $qtdOmissaoSped + $qtdOmissaoXml + $qtdSerieDivergente)
$totalExcecoes    = $qtdDivergenciaValor + $qtdOmissaoSped + $qtdOmissaoXml + $qtdSerieDivergente

Clear-Host
Write-Host ""
Write-Host " ============================================"
Write-Host "   MASSA DE TESTES GERADA COM SUCESSO"
Write-Host " ============================================"
Write-Host "   CNPJ AUDITADO            : $cnpjAuditado"
Write-Host "   ARQUIVOS XML GERADOS     : $totalXmlsGerados"
Write-Host "   REGISTROS SPED GERADOS   : $totalSpedGerados"
Write-Host " --------------------------------------------"
Write-Host "   NOTAS CONFORMES (OK)     : $qtdConforme"
Write-Host "   FILA DE REVISAO (ERROS)  : $totalExcecoes"
Write-Host "     - Divergencia de Valor : $qtdDivergenciaValor"
Write-Host "     - Omissoes no SPED     : $qtdOmissaoSped"
Write-Host "     - Omissoes no XML      : $qtdOmissaoXml"
Write-Host "     - Series Divergentes   : $qtdSerieDivergente"
Write-Host " ============================================"
Write-Host ""
Write-Host "  Agora rode o Munin.bat."
Write-Host "  Ele vai detectar automaticamente que os dados"
Write-Host "  sao de teste e perguntar se voce quer rodar"
Write-Host "  em modo DEMO (sem precisar configurar nada)."
Write-Host ""
