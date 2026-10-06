@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion
title MUNIN - Ymir

:: ============================================================
:: YMIR.bat
:: Orquestrador do MUNIN
:: ============================================================

set "BASE=%~dp0"
set "PASTA_XML=%BASE%01_XML_Fornecedores"
set "PASTA_SPED=%BASE%02_Dados_ERP"
set "CONFIG=%BASE%config.txt"
set "MUNIN_XLSX=%BASE%MUNIN.xlsx"
set "TEMPLATE_LOCAL=%BASE%MUNIN_template.xlsx"
set "MARKER_DEMO=%BASE%.huginn_demo"
set "SENHA_ARQ=%BASE%.senha.txt"

set "VERSAO_MUNIN=0.1.0"
set "TOLERANCIA_VALOR=0.01"

set "CNPJ_ATUAL="
set "TIPO_OPERACAO="
set "DATA_INICIO="
set "DATA_FIM="

set "DEMO_ATIVO=0"
set "DEMO_CNPJ="
set "DEMO_DT_INI="
set "DEMO_DT_FIM="

set "MODO_CONFIG=0"
if /I "%~1"=="/CONFIG" set "MODO_CONFIG=1"

if not exist "%PASTA_XML%" mkdir "%PASTA_XML%"
if not exist "%PASTA_SPED%" mkdir "%PASTA_SPED%"

echo.
echo  ============================================
echo   MUNIN - Motor de Unificacao, Navegacao
echo   e Inspecao de Notas
echo  ============================================
echo.

:: ---------- Senha das abas: le de .senha.txt ou gera nova ----------
set "SENHA_CORVOS="
if exist "%SENHA_ARQ%" (
    set /p SENHA_CORVOS=<"%SENHA_ARQ%"
)
if not defined SENHA_CORVOS (
    set "SENHA_CORVOS=Corvos%RANDOM%%RANDOM%%RANDOM%"
    > "%SENHA_ARQ%" echo !SENHA_CORVOS!
    echo  [i] Senha de protecao das abas gerada automaticamente.
    echo      Arquivo: %SENHA_ARQ%
    echo      Guarde este arquivo - ele protege o MUNIN.xlsx.
    echo.
)

if not exist "%MUNIN_XLSX%" (
    if not exist "%TEMPLATE_LOCAL%" (
        echo [ERRO] MUNIN.xlsx nao encontrado em: "%TEMPLATE_LOCAL%"
        pause
        exit /b 1
    )
    copy /Y "%TEMPLATE_LOCAL%" "%MUNIN_XLSX%" >nul
)

:: ============================================================
:: DETECCAO DE MODO DEMO (marker gerado pelo HUGINN)
:: ============================================================
if not exist "%MARKER_DEMO%" goto SEM_DEMO
if /I "%~1"=="/CONFIG" goto SEM_DEMO

for /f "usebackq tokens=1,* delims==" %%A in ("%MARKER_DEMO%") do (
    if /I "%%A"=="CNPJ" set "DEMO_CNPJ=%%B"
    if /I "%%A"=="DATA_INICIO" set "DEMO_DT_INI=%%B"
    if /I "%%A"=="DATA_FIM" set "DEMO_DT_FIM=%%B"
)

if not defined DEMO_CNPJ goto SEM_DEMO

echo.
echo  ============================================
echo   MUNIN - PROTOCOLO DEMO DETECTADO
echo  ============================================
echo.
echo   Os arquivos abaixo foram gerados pelo HUGINN,
echo   o gerador de massa de testes:
echo.
echo     01_XML_Fornecedores\
echo     02_Dados_ERP\
echo.
echo   CNPJ de teste  : !DEMO_CNPJ!
echo   Periodo        : !DEMO_DT_INI! a !DEMO_DT_FIM!
echo.
echo  --------------------------------------------
echo   Este e um protocolo DEMO?
echo.
echo    [S] Sim - rodar em modo demo com esses dados
echo    [N] Nao - configurar manualmente
echo  --------------------------------------------
echo.

set "RESP="
set /p "RESP= Sua escolha [S/N] (Enter = Sim): "
if not defined RESP set "RESP=S"

if /I "!RESP!"=="S"   goto DEMO_SIM
if /I "!RESP!"=="SIM" goto DEMO_SIM
if /I "!RESP!"=="Y"   goto DEMO_SIM
goto DEMO_NAO

:DEMO_SIM
set "DEMO_ATIVO=1"
del /q "%MARKER_DEMO%" >nul 2>&1
echo.
echo   Modo DEMO ativado. Usando CNPJ !DEMO_CNPJ!.
echo.
goto SEM_DEMO

:DEMO_NAO
del /q "%MARKER_DEMO%" >nul 2>&1
echo.
echo   Modo DEMO ignorado. Iniciando configuracao manual.
echo.
goto SEM_DEMO

:SEM_DEMO

:: ============================================================
:: LEITURA DA CONFIGURACAO
:: ============================================================
if exist "%CONFIG%" (
    for /f "usebackq tokens=1,* delims==" %%A in ("%CONFIG%") do (
        if /I "%%A"=="CNPJ_AUDITADO" set "CNPJ_ATUAL=%%B"
        if /I "%%A"=="TIPO_OPERACAO" set "TIPO_OPERACAO=%%B"
        if /I "%%A"=="DATA_INICIO" set "DATA_INICIO=%%B"
        if /I "%%A"=="DATA_FIM" set "DATA_FIM=%%B"
        if /I "%%A"=="TOLERANCIA_VALOR" set "TOLERANCIA_VALOR=%%B"
        if /I "%%A"=="VERSAO_MUNIN" set "VERSAO_MUNIN=%%B"
    )
)

if "!DEMO_ATIVO!"=="1" (
    set "CNPJ_ATUAL=!DEMO_CNPJ!"
    set "TIPO_OPERACAO=3"
    set "DATA_INICIO=!DEMO_DT_INI!"
    set "DATA_FIM=!DEMO_DT_FIM!"
    set "MODO_CONFIG=0"
) else (
    if not defined CNPJ_ATUAL set "MODO_CONFIG=1"
    if not defined TIPO_OPERACAO set "MODO_CONFIG=1"
    if not defined DATA_INICIO set "MODO_CONFIG=1"
    if not defined DATA_FIM set "MODO_CONFIG=1"
)

if "!MODO_CONFIG!"=="1" goto CONFIGURAR

goto CONFIGURACAO_OK

:CONFIGURAR
cls
echo.
echo  ============================================
echo   MUNIN - CONFIGURACAO DA AUDITORIA
echo  ============================================
echo.

:EDITAR_CNPJ
set "CNPJ_INPUT="
if defined CNPJ_ATUAL (
    set /p "CNPJ_INPUT= CNPJ auditado [!CNPJ_ATUAL!]: "
    if not defined CNPJ_INPUT set "CNPJ_INPUT=!CNPJ_ATUAL!"
) else (
    set /p "CNPJ_INPUT= CNPJ auditado: "
)

set "VALIDADOR=%TEMP%\ymir_valida_cnpj_%RANDOM%.ps1"
> "%VALIDADOR%" echo param([string]$c)
>> "%VALIDADOR%" echo $ErrorActionPreference = 'Stop'
>> "%VALIDADOR%" echo $s = ($c -replace '\W','').ToUpper()
>> "%VALIDADOR%" echo if ($s.Length -ne 14) { exit 1 }
>> "%VALIDADOR%" echo $b = $s.Substring(0,12)
>> "%VALIDADOR%" echo $dv = $s.Substring(12,2)
>> "%VALIDADOR%" echo if (($b -replace '[0-9A-Z]','') -ne '') { exit 1 }
>> "%VALIDADOR%" echo if (($dv -replace '[0-9]','') -ne '') { exit 1 }
>> "%VALIDADOR%" echo $p = 5,4,3,2,9,8,7,6,5,4,3,2
>> "%VALIDADOR%" echo $t = 0
>> "%VALIDADOR%" echo for ($i = 0; $i -lt 12; $i++) { $t += (([int][char]$b[$i]) - 48) * $p[$i] }
>> "%VALIDADOR%" echo $r = $t %% 11
>> "%VALIDADOR%" echo if ($r -lt 2) { $d1 = 0 } else { $d1 = 11 - $r }
>> "%VALIDADOR%" echo $p = 6,5,4,3,2,9,8,7,6,5,4,3,2
>> "%VALIDADOR%" echo $t = 0
>> "%VALIDADOR%" echo for ($i = 0; $i -lt 12; $i++) { $t += (([int][char]$b[$i]) - 48) * $p[$i] }
>> "%VALIDADOR%" echo $t += $d1 * 2
>> "%VALIDADOR%" echo $r = $t %% 11
>> "%VALIDADOR%" echo if ($r -lt 2) { $d2 = 0 } else { $d2 = 11 - $r }
>> "%VALIDADOR%" echo if (([string]$d1 + [string]$d2) -eq $dv) { Write-Output $s }

set "CNPJ_VALIDO="
for /f "delims=" %%A in ('powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%VALIDADOR%" "!CNPJ_INPUT!" 2^>nul') do set "CNPJ_VALIDO=%%A"
del /q "%VALIDADOR%" >nul 2>&1

if not defined CNPJ_VALIDO (
    echo [ERRO] CNPJ invalido - use 14 caracteres alfanumericos mais 2 digitos verificadores.
    set "CNPJ_ATUAL="
    timeout /t 2 >nul
    goto EDITAR_CNPJ
)
set "CNPJ_ATUAL=!CNPJ_VALIDO!"

:EDITAR_TIPO
echo.
echo  Tipo de nota: 1=Entrada, 2=Saida, 3=Geral
set "TIPO_INPUT="
set /p "TIPO_INPUT= Informe o tipo [!TIPO_OPERACAO!]: "
if not defined TIPO_INPUT set "TIPO_INPUT=!TIPO_OPERACAO!"

if "!TIPO_INPUT!"=="1" set "TIPO_OPERACAO=1" & goto TIPO_OK
if "!TIPO_INPUT!"=="2" set "TIPO_OPERACAO=2" & goto TIPO_OK
if "!TIPO_INPUT!"=="3" set "TIPO_OPERACAO=3" & goto TIPO_OK
goto EDITAR_TIPO

:TIPO_OK

:EDITAR_DATA_INICIO
echo.
set "DATA_INPUT="
set /p "DATA_INPUT= Data inicial [!DATA_INICIO!]: "
if not defined DATA_INPUT set "DATA_INPUT=!DATA_INICIO!"

set "DATA_VALIDADA="
for /f "delims=" %%A in ('powershell.exe -NoProfile -Command "& { param($v); $fmts = @('dd/MM/yyyy','yyyy-MM-dd','dd-MM-yyyy','yyyy/MM/dd','ddMMyyyy','yyyyMMdd'); $d = [datetime]::MinValue; foreach($f in $fmts){ if([datetime]::TryParseExact($v.Trim(), $f, [cultureinfo]::InvariantCulture, [datetimestyles]::None, [ref]$d)){ Write-Output $d.ToString('dd/MM/yyyy'); break } } }" "!DATA_INPUT!"') do (
    set "DATA_VALIDADA=%%A"
)
if not defined DATA_VALIDADA (
    echo [ERRO] Data inicial invalida.
    goto EDITAR_DATA_INICIO
)
set "DATA_INICIO=!DATA_VALIDADA!"

:EDITAR_DATA_FIM
echo.
set "DATA_INPUT="
set /p "DATA_INPUT= Data final [!DATA_FIM!]: "
if not defined DATA_INPUT set "DATA_INPUT=!DATA_FIM!"

set "DATA_VALIDADA="
for /f "delims=" %%A in ('powershell.exe -NoProfile -Command "& { param($v); $fmts = @('dd/MM/yyyy','yyyy-MM-dd','dd-MM-yyyy','yyyy/MM/dd','ddMMyyyy','yyyyMMdd'); $d = [datetime]::MinValue; foreach($f in $fmts){ if([datetime]::TryParseExact($v.Trim(), $f, [cultureinfo]::InvariantCulture, [datetimestyles]::None, [ref]$d)){ Write-Output $d.ToString('dd/MM/yyyy'); break } } }" "!DATA_INPUT!"') do (
    set "DATA_VALIDADA=%%A"
)
if not defined DATA_VALIDADA (
    echo [ERRO] Data final invalida.
    goto EDITAR_DATA_FIM
)
set "DATA_FIM=!DATA_VALIDADA!"

:CONFIGURACAO_OK

:VERIFICA_ARQUIVOS
set "QTD_XML=0"
for %%F in ("%PASTA_XML%\*.xml") do if exist "%%~fF" set /a QTD_XML+=1

set "QTD_SPED=0"
set "ARQUIVO_SPED="
for %%F in ("%PASTA_SPED%\*.txt") do (
    if exist "%%~fF" (
        set /a QTD_SPED+=1
        set "ARQUIVO_SPED=%%~fF"
    )
)

if !QTD_SPED! GTR 1 (
    echo [ERRO] Foram encontrados !QTD_SPED! arquivos SPED TXT.
    pause
    exit /b 1
)
if !QTD_XML! GTR 0 if !QTD_SPED! EQU 1 goto ARQUIVOS_OK

timeout /t 10 >nul
goto VERIFICA_ARQUIVOS

:ARQUIVOS_OK

> "%CONFIG%" echo CNPJ_AUDITADO=!CNPJ_ATUAL!
>> "%CONFIG%" echo PASTA_XML=%PASTA_XML%
>> "%CONFIG%" echo ARQUIVO_SPED=!ARQUIVO_SPED!
>> "%CONFIG%" echo TOLERANCIA_VALOR=%TOLERANCIA_VALOR%
>> "%CONFIG%" echo VERSAO_MUNIN=%VERSAO_MUNIN%
>> "%CONFIG%" echo TIPO_OPERACAO=!TIPO_OPERACAO!
>> "%CONFIG%" echo DATA_INICIO=!DATA_INICIO!
>> "%CONFIG%" echo DATA_FIM=!DATA_FIM!

set "PS1=%TEMP%\ymir_update_%RANDOM%_%RANDOM%.ps1"

> "%PS1%" echo param(
>> "%PS1%" echo     [string]$XlsxPath,
>> "%PS1%" echo     [string]$Cnpj,
>> "%PS1%" echo     [string]$PastaXml,
>> "%PS1%" echo     [string]$ArquivoSped,
>> "%PS1%" echo     [string]$Tolerancia,
>> "%PS1%" echo     [string]$Versao,
>> "%PS1%" echo     [string]$TipoOp,
>> "%PS1%" echo     [string]$DtIni,
>> "%PS1%" echo     [string]$DtFim,
>> "%PS1%" echo     [string]$SenhaCorvos
>> "%PS1%" echo )
>> "%PS1%" echo $ErrorActionPreference = 'Stop'
>> "%PS1%" echo $excel = $null
>> "%PS1%" echo $wb = $null
>> "%PS1%" echo $abasProtegidas = @()
>> "%PS1%" echo try {
>> "%PS1%" echo     Write-Host '[1/7] Abrindo Excel...'
>> "%PS1%" echo     $excel = New-Object -ComObject Excel.Application
>> "%PS1%" echo     $excel.Visible = $true
>> "%PS1%" echo     $excel.DisplayAlerts = $false
>> "%PS1%" echo
>> "%PS1%" echo     Write-Host '[2/7] Criando backup do workbook...'
>> "%PS1%" echo     if (Test-Path -LiteralPath $XlsxPath) {
>> "%PS1%" echo         Copy-Item -LiteralPath $XlsxPath -Destination "$XlsxPath.bak" -Force
>> "%PS1%" echo     }
>> "%PS1%" echo
>> "%PS1%" echo     $wb = $excel.Workbooks.Open($XlsxPath, 0, $false)
>> "%PS1%" echo
>> "%PS1%" echo     Write-Host '[2.5/7] Verificando protecao das abas...'
>> "%PS1%" echo     foreach ($ws in $wb.Worksheets) {
>> "%PS1%" echo         if ($ws.ProtectContents) {
>> "%PS1%" echo             $abasProtegidas += $ws.Name
>> "%PS1%" echo             try { $ws.Unprotect($SenhaCorvos) } catch { try { $ws.Unprotect() } catch {} }
>> "%PS1%" echo         }
>> "%PS1%" echo     }
>> "%PS1%" echo
>> "%PS1%" echo     Write-Host '[3/7] Localizando configuracao...'
>> "%PS1%" echo     $cfg = $null
>> "%PS1%" echo     foreach ($ws in $wb.Worksheets) {
>> "%PS1%" echo         if ($ws.Name -eq 'CONFIGURACAO' -or $ws.Name -eq '10_CONFIG') { $cfg = $ws; break }
>> "%PS1%" echo     }
>> "%PS1%" echo     if ($null -eq $cfg) { throw 'Aba CONFIGURACAO/10_CONFIG nao encontrada.' }
>> "%PS1%" echo
>> "%PS1%" echo     $tbl = $null
>> "%PS1%" echo     foreach ($lo in $cfg.ListObjects) {
>> "%PS1%" echo         if ($lo.Name -eq 'tblConfig') { $tbl = $lo; break }
>> "%PS1%" echo     }
>> "%PS1%" echo     if ($null -eq $tbl) { throw 'Tabela tblConfig nao encontrada.' }
>> "%PS1%" echo
>> "%PS1%" echo     function Set-ConfigValor([string]$Chave, [string]$Valor) {
>> "%PS1%" echo         $data = $tbl.DataBodyRange
>> "%PS1%" echo         if ($null -ne $data) {
>> "%PS1%" echo             for ($i = 1; $i -le $tbl.ListRows.Count; $i++) {
>> "%PS1%" echo                 $celChave = [string]$data.Cells.Item($i, 1).Text
>> "%PS1%" echo                 if ($celChave.Trim() -eq $Chave) {
>> "%PS1%" echo                     $data.Cells.Item($i, 2).Value2 = $Valor
>> "%PS1%" echo                     return
>> "%PS1%" echo                 }
>> "%PS1%" echo             }
>> "%PS1%" echo         }
>> "%PS1%" echo         $nova = $tbl.ListRows.Add()
>> "%PS1%" echo         $nova.Range.Cells.Item(1,1).Value2 = $Chave
>> "%PS1%" echo         $nova.Range.Cells.Item(1,2).Value2 = $Valor
>> "%PS1%" echo     }
>> "%PS1%" echo
>> "%PS1%" echo     Write-Host '[4/7] Gravando parametros na tblConfig...'
>> "%PS1%" echo     Set-ConfigValor 'CNPJ_AUDITADO' $Cnpj
>> "%PS1%" echo     Set-ConfigValor 'PASTA_XML' $PastaXml
>> "%PS1%" echo     Set-ConfigValor 'ARQUIVO_SPED' $ArquivoSped
>> "%PS1%" echo     Set-ConfigValor 'TOLERANCIA_VALOR' $Tolerancia
>> "%PS1%" echo     Set-ConfigValor 'VERSAO_MUNIN' $Versao
>> "%PS1%" echo     Set-ConfigValor 'TIPO_OPERACAO' $TipoOp
>> "%PS1%" echo     Set-ConfigValor 'DATA_INICIO' $DtIni
>> "%PS1%" echo     Set-ConfigValor 'DATA_FIM' $DtFim
>> "%PS1%" echo
>> "%PS1%" echo     Write-Host '[5/7] Verificando consultas Power Query...'
>> "%PS1%" echo     $wb.Save()
>> "%PS1%" echo
>> "%PS1%" echo     Write-Host '[6/7] Atualizando consultas...'
>> "%PS1%" echo     $wb.RefreshAll()
>> "%PS1%" echo
>> "%PS1%" echo     $limite = (Get-Date).AddSeconds(120)
>> "%PS1%" echo     do {
>> "%PS1%" echo         Start-Sleep -Seconds 2
>> "%PS1%" echo         try { $excel.CalculateUntilAsyncQueriesDone() } catch { }
>> "%PS1%" echo     } while ((Get-Date) -lt $limite)
>> "%PS1%" echo
>> "%PS1%" echo     Write-Host '[7/7] Finalizando...'
>> "%PS1%" echo
>> "%PS1%" echo     Write-Host '      Re-protegendo abas...'
>> "%PS1%" echo     foreach ($nomeAba in $abasProtegidas) {
>> "%PS1%" echo         try { $wb.Worksheets.Item($nomeAba).Protect($SenhaCorvos) } catch {}
>> "%PS1%" echo     }
>> "%PS1%" echo
>> "%PS1%" echo     $wb.Save()
>> "%PS1%" echo
>> "%PS1%" echo     foreach ($ws in $wb.Worksheets) {
>> "%PS1%" echo         if ($ws.Name -eq 'PAINEL' -or $ws.Name -eq '01_PAINEL') { $ws.Activate(); break }
>> "%PS1%" echo     }
>> "%PS1%" echo     $excel.DisplayAlerts = $true
>> "%PS1%" echo     Write-Host '[OK] Auditoria atualizada com sucesso.'
>> "%PS1%" echo } catch {
>> "%PS1%" echo     Write-Host ('[ERRO] ' + $_.Exception.Message)
>> "%PS1%" echo     try { if ($null -ne $wb) { $wb.Close($false) } } catch {}
>> "%PS1%" echo     try { if ($null -ne $excel) { $excel.Quit() } } catch {}
>> "%PS1%" echo     exit 1
>> "%PS1%" echo }

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%" ^
    -XlsxPath "%MUNIN_XLSX%" ^
    -Cnpj "!CNPJ_ATUAL!" ^
    -PastaXml "%PASTA_XML%" ^
    -ArquivoSped "!ARQUIVO_SPED!" ^
    -Tolerancia "%TOLERANCIA_VALOR%" ^
    -Versao "%VERSAO_MUNIN%" ^
    -TipoOp "!TIPO_OPERACAO!" ^
    -DtIni "!DATA_INICIO!" ^
    -DtFim "!DATA_FIM!" ^
    -SenhaCorvos "!SENHA_CORVOS!"

set "PS_STATUS=%ERRORLEVEL%"
del /q "%PS1%" >nul 2>&1

if not "%PS_STATUS%"=="0" (
    echo [ERRO] Falha ao atualizar Excel.
    pause
    exit /b %PS_STATUS%
)

echo.
echo ============================================
echo  MUNIN - ATUALIZADO COM SUCESSO
echo ============================================
echo.
pause
endlocal
exit /b 0
