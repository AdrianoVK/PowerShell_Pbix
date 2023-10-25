# O módulo MicrosoftPowerBIMgmt abaixo deve ser instalado apenas na primeira vez que executa esse script, depois não há 
# necidades de chamar esse modulo depois. Basta comentar ou excluir depois da primeira chamada.
####  Install-Module -Name MicrosoftPowerBIMgmt


#Lembre-se que o powerShell só conseguimos executar utilizando ele como Admin, sem isso o script não funciona

Connect-PowerBIServiceAccount 

### ESTES PBIX FOI DESENVOLVIDO PARA SEGURADORAS E CORRETORAS. NESTE CASO AS VARIAVEIS SEGUIRÁ COMO ORIGINAL DESSA APLICACAO ###
#criando as variaveis para pegar os pbix de Corretoras
$corretora = Get-PowerBIWorkspace -Name 'WORKSPACE ONDE O PBIX ESTA ALOCADO' 
$relatorios_corr = Get-PowerBIReport -Workspace $corretoras
$caminho = "CAMINHO ONDE VOCÊ QUER FAZER O DOWNLOAD"
#fazendo o iteraveil para saber em que posição de download estamos
$a = 0
$qtdrelatorios = $relatorios_corr.Count

#Baixando os PBIX Workspace Corretoras
foreach ($relatorio in $relatorios_corr){
    $a++
    $nomeArquivo = $relatorio.name + ".pbix"
    $arquivo = Join-Path $caminho $nomeArquivo
    Write-Host "Downloading $a | $qtdrelatorios - $nomeArquivo"
    Export-PowerBIReport -Id $relatorio.ID -OutFile $arquivo
}

#criando as variaveis para pegar os pbix de Seguradoras
$seguradora = Get-PowerBIWorkspace -Name 'WORKSPACE ONDE O PBIX ESTA ALOCADO' 
$relatorios_seg = Get-PowerBIReport -Workspace $seguradora
$caminho = "CAMINHO ONDE VOCÊ QUER FAZER O DOWNLOAD"
#fazendo o iteraveil para saber em que posição de download estamos
$b = 0
$qtdrelatorios = $relatorios_seg.Count

# Baixando os PBIX Workspace Seguradoras
foreach ($relatorio in $relatorios_seg){
    $b++
    $nomeArquivo = $relatorio.name + ".pbix"
    $arquivo = Join-Path $caminho $nomeArquivo
    Write-Host "Downloading $b | $qtdrelatorios - $nomeArquivo"
    Export-PowerBIReport -Id $relatorio.ID -OutFile $arquivo
}

#criando as variaveis para pegar os pbix VISÃO GERENCIAL
$RMS = Get-PowerBIWorkspace -Name 'WORKSPACE ONDE O PBIX ESTA ALOCADO'
$RMS_pbi = Get-PowerBIReport -Workspace $RMS
$caminho = "CAMINHO ONDE VOCÊ QUER FAZER O DOWNLOAD"
#fazendo o iteraveil para saber em que posição de download estamos
$c = 0
$qtdrelatorios = $RMS_pbi.Count

# Baixando os PBIX Workspace VISÃO GERENCIAL
foreach ($relatorio in $RMS_pbi){
    $c++
    $nomeArquivo = $relatorio.name + ".pbix"
    $arquivo = Join-Path $caminho $nomeArquivo
    Write-Host "Downloading $c | $qtdrelatorios - $nomeArquivo"
    Export-PowerBIReport -Id $relatorio.ID -OutFile $arquivo
}

#criando as variaveis para pegar os pbix de VISÃO GERENCIAL 2
$macros_seguradoras = Get-PowerBIWorkspace -Name 'WORKSPACE ONDE O PBIX ESTA ALOCADO'
$macro_seg = Get-PowerBIReport -Workspace $macros_seguradoras
$caminho = "CAMINHO ONDE VOCÊ QUER FAZER O DOWNLOAD"
#fazendo o iteraveil para saber em que posição de download estamos
$d = 0
$qtdrelatorios = $macro_seg.Count

# Baixando os PBIX Workspace VISÃO GERENCIAL 2
foreach ($relatorio in $macro_seg){
    $d++
    $nomeArquivo = $relatorio.name + ".pbix"
    $arquivo = Join-Path $caminho $nomeArquivo
    Write-Host "Downloading $d | $qtdrelatorios - $nomeArquivo"
    Export-PowerBIReport -Id $relatorio.ID -OutFile $arquivo
}

#criando as variaveis para pegar os pbix de VISÃO GERENCIAL 3
$macros_corretoras = Get-PowerBIWorkspace -Name 'WORKSPACE ONDE O PBIX ESTA ALOCADO'
$macro_corr = Get-PowerBIReport -Workspace $macros_corretoras
$caminho = "CAMINHO ONDE VOCÊ QUER FAZER O DOWNLOAD"
#fazendo o iteraveil para saber em que posição de download estamos
$e = 0
$qtdrelatorios = $macros_corr.Count

# Baixando os PBIX Workspace VISÃO GERENCIAL 3
foreach ($relatorio in $macro_corr){
    $e++
    $nomeArquivo = $relatorio.name + ".pbix"
    $arquivo = Join-Path $caminho $nomeArquivo
    Write-Host "Downloading $e | $qtdrelatorios - $nomeArquivo"
    Export-PowerBIReport -Id $relatorio.ID -OutFile $arquivo
}
