$origem = "C:\Users\iNFORMARK Loja\Desktop\iphone-inteligencia\Bot"
$docs = Join-Path $origem "docs"

New-Item -ItemType Directory -Force -Path $docs | Out-Null

$arquivoRelatorio = Get-ChildItem -Path $origem -Filter "relatorio_menor_preco_*.csv" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

$arquivoPrecos = Join-Path $origem "precos.csv"

function Nova-TabelaHtml {
    param (
        [array]$Dados,
        [string]$IdTabela
    )

    if (-not $Dados -or $Dados.Count -eq 0) {
        return "<p>Nenhum dado encontrado.</p>"
    }

    $colunas = $Dados[0].PSObject.Properties.Name

    $thead = ""
    foreach ($col in $colunas) {
        $thead += "<th>$col</th>`n"
    }

    $tbody = ""
    foreach ($linha in $Dados) {
        $tbody += "<tr>`n"
        foreach ($col in $colunas) {
            $valor = $linha.$col
            if ($null -eq $valor) { $valor = "" }
            $tbody += "<td>$valor</td>`n"
        }
        $tbody += "</tr>`n"
    }

    return @"
<div class="table-wrap">
    <table id="$IdTabela">
        <thead>
            <tr>
$thead
            </tr>
        </thead>
        <tbody>
$tbody
        </tbody>
    </table>
</div>
"@
}

$dadosRelatorio = @()
$dadosPrecos = @()

if ($arquivoRelatorio) {
    $dadosRelatorio = Import-Csv $arquivoRelatorio.FullName
}

if (Test-Path $arquivoPrecos) {
    $dadosPrecos = Import-Csv $arquivoPrecos
}

$tabelaRelatorio = Nova-TabelaHtml -Dados $dadosRelatorio -IdTabela "tabelaRelatorio"
$tabelaPrecos = Nova-TabelaHtml -Dados $dadosPrecos -IdTabela "tabelaPrecos"

$atualizadoEm = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
$nomeArquivoRelatorio = if ($arquivoRelatorio) { $arquivoRelatorio.Name } else { "Nenhum relatório encontrado" }
$nomeArquivoPrecos = if (Test-Path $arquivoPrecos) { "precos.csv" } else { "precos.csv não encontrado" }

$html = @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relat&oacute;rio Informark</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background: #f3f4f6;
            margin: 0;
            padding: 20px;
            color: #111827;
        }
        .container {
            max-width: 1300px;
            margin: 0 auto;
        }
        .card {
            background: #ffffff;
            border-radius: 16px;
            padding: 20px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        }
        h1 {
            margin-top: 0;
            margin-bottom: 10px;
        }
        .meta {
            color: #6b7280;
            margin-bottom: 8px;
        }
        .tabs {
            display: flex;
            gap: 10px;
            margin: 20px 0 15px 0;
            flex-wrap: wrap;
        }
        .tab-btn {
            border: none;
            background: #e5e7eb;
            color: #111827;
            padding: 10px 16px;
            border-radius: 10px;
            cursor: pointer;
            font-size: 15px;
            font-weight: bold;
        }
        .tab-btn.active {
            background: #111827;
            color: white;
        }
        .tab-content {
            display: none;
        }
        .tab-content.active {
            display: block;
        }
        input {
            width: 100%;
            padding: 12px;
            font-size: 16px;
            border: 1px solid #d1d5db;
            border-radius: 10px;
            box-sizing: border-box;
            margin: 15px 0;
        }
        .table-wrap {
            overflow-x: auto;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            min-width: 900px;
            background: white;
        }
        th, td {
            padding: 10px 12px;
            border-bottom: 1px solid #e5e7eb;
            text-align: left;
            white-space: nowrap;
        }
        th {
            background: #111827;
            color: white;
            position: sticky;
            top: 0;
        }
        tr:hover td {
            background: #f9fafb;
        }
        .subtitulo {
            margin-top: 10px;
            margin-bottom: 8px;
            font-weight: bold;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="card">
            <h1>Relat&oacute;rio Informark</h1>
            <div class="meta">&Uacute;ltima atualiza&ccedil;&atilde;o: $atualizadoEm</div>

            <div class="tabs">
                <button class="tab-btn active" onclick="abrirAba('abaRelatorio', this)">Menor Pre&ccedil;o</button>
                <button class="tab-btn" onclick="abrirAba('abaPrecos', this)">Planilha de Pre&ccedil;os</button>
            </div>

            <div id="abaRelatorio" class="tab-content active">
                <div class="subtitulo">Arquivo base: $nomeArquivoRelatorio</div>
                <input type="text" id="buscaRelatorio" placeholder="Buscar no relat&oacute;rio de menor pre&ccedil;o...">
                $tabelaRelatorio
            </div>

            <div id="abaPrecos" class="tab-content">
                <div class="subtitulo">Arquivo base: $nomeArquivoPrecos</div>
                <input type="text" id="buscaPrecos" placeholder="Buscar na planilha de pre&ccedil;os...">
                $tabelaPrecos
            </div>
        </div>
    </div>

    <script>
        function abrirAba(idAba, botao) {
            document.querySelectorAll('.tab-content').forEach(function(aba) {
                aba.classList.remove('active');
            });

            document.querySelectorAll('.tab-btn').forEach(function(btn) {
                btn.classList.remove('active');
            });

            document.getElementById(idAba).classList.add('active');
            botao.classList.add('active');
        }

        function ativarBusca(inputId, tabelaId) {
            const busca = document.getElementById(inputId);
            const linhas = document.querySelectorAll('#' + tabelaId + ' tbody tr');

            if (!busca) return;

            busca.addEventListener('input', function () {
                const termo = this.value.toLowerCase();

                linhas.forEach(function(linha) {
                    const texto = linha.innerText.toLowerCase();
                    linha.style.display = texto.includes(termo) ? '' : 'none';
                });
            });
        }

        ativarBusca('buscaRelatorio', 'tabelaRelatorio');
        ativarBusca('buscaPrecos', 'tabelaPrecos');
    </script>
</body>
</html>
"@

$destino = Join-Path $docs "index.html"
[System.IO.File]::WriteAllText($destino, $html, [System.Text.UTF8Encoding]::new($false))

Write-Host "HTML gerado em docs\index.html com duas abas"