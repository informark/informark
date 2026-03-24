$origem = "C:\Users\iNFORMARK Loja\Desktop\iphone-inteligencia\Bot"
$docs = Join-Path $origem "docs"

New-Item -ItemType Directory -Force -Path $docs | Out-Null

$arquivoRelatorio = Get-ChildItem -Path $origem -Filter "relatorio_menor_preco_*.csv" -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

$arquivoPrecos = Join-Path $origem "precos.csv"
$arquivoPrecoDia = Join-Path $origem "preco_dia.csv"

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
        $colEscapado = [System.Net.WebUtility]::HtmlEncode([string]$col)
        $thead += "<th>$colEscapado</th>`n"
    }

    $json = $Dados | ConvertTo-Json -Depth 5 -Compress
    $jsonSeguro = $json -replace '</script>', '<\/script>'

    return @"
<div class="table-wrap">
    <table id="$IdTabela">
        <thead>
            <tr>
$thead
            </tr>
        </thead>
        <tbody></tbody>
    </table>
</div>
<div class="paginacao" id="${IdTabela}_paginacao"></div>
<script type="application/json" id="${IdTabela}_data">$jsonSeguro</script>
"@
}

function Importar-XlsxComoObjetos {
    param (
        [string]$CaminhoArquivo
    )

    $resultado = @()

    if (-not (Test-Path $CaminhoArquivo)) {
        return $resultado
    }

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $usedRange = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($CaminhoArquivo)
        $worksheet = $workbook.Worksheets.Item(1)
        $usedRange = $worksheet.UsedRange

        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count

        if ($rowCount -lt 2 -or $colCount -lt 1) {
            return @()
        }

        $headers = @()
        for ($col = 1; $col -le $colCount; $col++) {
            $headerText = [string]$usedRange.Cells.Item(1, $col).Text
            if ([string]::IsNullOrWhiteSpace($headerText)) {
                $headerText = "Coluna$col"
            }
            $headers += $headerText.Trim()
        }

        for ($row = 2; $row -le $rowCount; $row++) {
            $obj = [ordered]@{}
            $temConteudo = $false

            for ($col = 1; $col -le $colCount; $col++) {
                $valor = [string]$usedRange.Cells.Item($row, $col).Text
                if (-not [string]::IsNullOrWhiteSpace($valor)) {
                    $temConteudo = $true
                }
                $obj[$headers[$col - 1]] = $valor
            }

            if ($temConteudo) {
                $resultado += [PSCustomObject]$obj
            }
        }
    }
    finally {
        if ($workbook) { $workbook.Close($false) | Out-Null }
        if ($excel) { $excel.Quit() | Out-Null }

        if ($usedRange) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) }
        if ($worksheet) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) }
        if ($workbook) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
        if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    return $resultado
}

$dadosRelatorio = @()
$dadosPrecos = @()
$dadosPrecoDia = @()

if ($arquivoRelatorio) {
    $dadosRelatorio = Import-Csv $arquivoRelatorio.FullName
}

if (Test-Path $arquivoPrecos) {
    $dadosPrecos = Import-Csv $arquivoPrecos
}

if (Test-Path $arquivoPrecoDia) {
    $dadosPrecoDia = Import-Csv $arquivoPrecoDia
}

$tabelaRelatorio = Nova-TabelaHtml -Dados $dadosRelatorio -IdTabela "tabelaRelatorio"
$tabelaPrecos = Nova-TabelaHtml -Dados $dadosPrecos -IdTabela "tabelaPrecos"
$tabelaPrecoDia = Nova-TabelaHtml -Dados $dadosPrecoDia -IdTabela "tabelaPrecoDia"

$atualizadoEm = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
$nomeArquivoRelatorio = if ($arquivoRelatorio) { $arquivoRelatorio.Name } else { "Nenhum relatório encontrado" }
$nomeArquivoPrecos = if (Test-Path $arquivoPrecos) { "precos.csv" } else { "precos.csv não encontrado" }
$nomeArquivoPrecoDia = if (Test-Path $arquivoPrecoDia) { "preco_dia.csv" } else { "preco_dia.csv não encontrado" }

$html = @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Informark Dashboard</title>
    <style>
        * { box-sizing: border-box; }

        body {
            font-family: Arial, sans-serif;
            background: #f3f4f6;
            margin: 0;
            padding: 20px;
            color: #111827;
        }

        .container {
            max-width: 1450px;
            margin: 0 auto;
        }

        .hero {
            background: linear-gradient(135deg, #0f172a, #1e293b);
            color: white;
            border-radius: 22px;
            padding: 28px;
            margin-bottom: 18px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.18);
        }

        .hero h1 {
            margin: 0 0 8px 0;
            font-size: 34px;
        }

        .hero p {
            margin: 0;
            color: #cbd5e1;
            font-size: 15px;
        }

        .pill {
            display: inline-block;
            padding: 6px 10px;
            border-radius: 999px;
            background: #e2e8f0;
            color: #0f172a;
            font-size: 12px;
            font-weight: bold;
            margin-top: 10px;
        }

        .stats {
            display: grid;
            grid-template-columns: repeat(5, minmax(160px, 1fr));
            gap: 14px;
            margin-bottom: 18px;
        }

        .stat-card {
            background: white;
            border-radius: 18px;
            padding: 18px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        }

        .stat-label {
            color: #6b7280;
            font-size: 13px;
            margin-bottom: 8px;
        }

        .stat-value {
            font-size: 24px;
            font-weight: bold;
            color: #111827;
            word-break: break-word;
        }

        .main-card {
            background: #ffffff;
            border-radius: 20px;
            padding: 22px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        }

        .meta {
            color: #6b7280;
            margin-bottom: 6px;
            font-size: 14px;
        }

        .tabs {
            display: flex;
            gap: 10px;
            margin: 20px 0 16px 0;
            flex-wrap: wrap;
        }

        .tab-btn {
            border: none;
            background: #e5e7eb;
            color: #111827;
            padding: 12px 18px;
            border-radius: 12px;
            cursor: pointer;
            font-size: 15px;
            font-weight: bold;
        }

        .tab-btn.active {
            background: #0f172a;
            color: white;
        }

        .tab-content {
            display: none;
        }

        .tab-content.active {
            display: block;
        }

        .subtitulo {
            margin-top: 6px;
            margin-bottom: 10px;
            font-weight: bold;
            font-size: 16px;
        }

        .filtros {
            display: grid;
            grid-template-columns: repeat(8, minmax(140px, 1fr));
            gap: 10px;
            margin: 15px 0 12px 0;
        }

        input, select {
            width: 100%;
            padding: 12px;
            font-size: 15px;
            border: 1px solid #d1d5db;
            border-radius: 12px;
            background: white;
            color: #111827;
        }

        .resumo {
            font-size: 14px;
            color: #4b5563;
            margin-bottom: 14px;
            font-weight: 600;
        }

        .table-wrap {
            overflow-x: auto;
            border-radius: 16px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            min-width: 1000px;
            background: white;
        }

        th, td {
            padding: 12px 14px;
            border-bottom: 1px solid #e5e7eb;
            text-align: left;
            white-space: nowrap;
        }

        th {
            background: #0f172a;
            color: white;
            position: sticky;
            top: 0;
            z-index: 1;
        }

        tr:hover td {
            background: #f8fafc;
        }

        .paginacao {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            align-items: center;
            margin-top: 14px;
        }

        .paginacao button {
            border: none;
            background: #0f172a;
            color: white;
            padding: 10px 14px;
            border-radius: 10px;
            cursor: pointer;
            font-size: 14px;
        }

        .paginacao button:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }

        .paginacao .info-pagina {
            font-size: 14px;
            color: #4b5563;
            font-weight: 600;
        }

        @media (max-width: 1200px) {
            .stats {
                grid-template-columns: repeat(3, minmax(140px, 1fr));
            }

            .filtros {
                grid-template-columns: repeat(2, 1fr);
            }
        }

        @media (max-width: 700px) {
            body {
                padding: 12px;
            }

            .hero {
                padding: 20px;
            }

            .hero h1 {
                font-size: 26px;
            }

            .main-card {
                padding: 16px;
            }

            .stats {
                grid-template-columns: 1fr 1fr;
            }

            .filtros {
                grid-template-columns: 1fr;
            }

            .stat-value {
                font-size: 20px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <section class="hero">
            <h1>Informark Dashboard</h1>
            <p>Painel de acompanhamento de pre&ccedil;os e relat&oacute;rios do bot.</p>
            <div class="pill">Atualizado em $atualizadoEm</div>
        </section>

        <section class="stats">
            <div class="stat-card">
                <div class="stat-label">Registros vis&iacute;veis</div>
                <div class="stat-value" id="statRegistros">0</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Modelos vis&iacute;veis</div>
                <div class="stat-value" id="statModelos">0</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Menor pre&ccedil;o vis&iacute;vel</div>
                <div class="stat-value" id="statMinPreco">-</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Maior pre&ccedil;o vis&iacute;vel</div>
                <div class="stat-value" id="statMaxPreco">-</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Aba atual</div>
                <div class="stat-value" id="statAba">Menor Pre&ccedil;o</div>
            </div>
        </section>

        <section class="main-card">
            <div class="meta">&Uacute;ltima atualiza&ccedil;&atilde;o: $atualizadoEm</div>

            <div class="tabs">
                <button class="tab-btn active" onclick="abrirAba('abaRelatorio', this, 'Menor Preço')">Menor Pre&ccedil;o</button>
                <button class="tab-btn" onclick="abrirAba('abaPrecos', this, 'Planilha de Preços')">Planilha de Pre&ccedil;os</button>
                <button class="tab-btn" onclick="abrirAba('abaPrecoDia', this, 'Preco do dia')">Preco do dia</button>
            </div>

            <div id="abaRelatorio" class="tab-content active">
                <div class="subtitulo">Arquivo base: $nomeArquivoRelatorio</div>

                <div class="filtros">
                    <input type="text" id="buscaRelatorio" placeholder="Buscar no relat&oacute;rio...">
                    <select id="produtoRelatorio"><option value="">Todos os produtos</option></select>
                    <select id="modeloRelatorio"><option value="">Todos os modelos</option></select>
                    <select id="gbRelatorio"><option value="">Todos os GB</option></select>
                    <select id="condicaoRelatorio"><option value="">Todas as condi&ccedil;&otilde;es</option></select>
                    <input type="number" id="precoMinRelatorio" placeholder="Pre&ccedil;o m&iacute;nimo">
                    <input type="number" id="precoMaxRelatorio" placeholder="Pre&ccedil;o m&aacute;ximo">
                    <select id="ordenacaoRelatorio">
                        <option value="">Ordena&ccedil;&atilde;o padr&atilde;o</option>
                        <option value="preco-asc">Pre&ccedil;o: menor para maior</option>
                        <option value="preco-desc">Pre&ccedil;o: maior para menor</option>
                    </select>
                </div>

                <div class="resumo" id="resumoRelatorio"></div>
                $tabelaRelatorio
            </div>

            <div id="abaPrecos" class="tab-content">
                <div class="subtitulo">Arquivo base: $nomeArquivoPrecos</div>

                <div class="filtros">
                    <input type="text" id="buscaPrecos" placeholder="Buscar na planilha de pre&ccedil;os...">
                    <select id="produtoPrecos"><option value="">Todos os produtos</option></select>
                    <select id="modeloPrecos"><option value="">Todos os modelos</option></select>
                    <select id="gbPrecos"><option value="">Todos os GB</option></select>
                    <select id="condicaoPrecos"><option value="">Todas as condi&ccedil;&otilde;es</option></select>
                    <input type="number" id="precoMinPrecos" placeholder="Pre&ccedil;o m&iacute;nimo">
                    <input type="number" id="precoMaxPrecos" placeholder="Pre&ccedil;o m&aacute;ximo">
                    <select id="ordenacaoPrecos">
                        <option value="">Ordena&ccedil;&atilde;o padr&atilde;o</option>
                        <option value="preco-asc">Pre&ccedil;o: menor para maior</option>
                        <option value="preco-desc">Pre&ccedil;o: maior para menor</option>
                    </select>
                </div>

                <div class="resumo" id="resumoPrecos"></div>
                $tabelaPrecos
            </div>

            <div id="abaPrecoDia" class="tab-content">
                <div class="subtitulo">Arquivo base: $nomeArquivoPrecoDia</div>

                <div class="filtros">
                    <input type="text" id="buscaPrecoDia" placeholder="Buscar na planilha preco do dia...">
                    <select id="produtoPrecoDia"><option value="">Todos os produtos</option></select>
                    <select id="modeloPrecoDia"><option value="">Todos os modelos</option></select>
                    <select id="gbPrecoDia"><option value="">Todos os GB</option></select>
                    <select id="condicaoPrecoDia"><option value="">Todas as condi&ccedil;&otilde;es</option></select>
                    <input type="number" id="precoMinPrecoDia" placeholder="Pre&ccedil;o m&iacute;nimo">
                    <input type="number" id="precoMaxPrecoDia" placeholder="Pre&ccedil;o m&aacute;ximo">
                    <select id="ordenacaoPrecoDia">
                        <option value="">Ordena&ccedil;&atilde;o padr&atilde;o</option>
                        <option value="preco-asc">Pre&ccedil;o: menor para maior</option>
                        <option value="preco-desc">Pre&ccedil;o: maior para menor</option>
                    </select>
                </div>

                <div class="resumo" id="resumoPrecoDia"></div>
                $tabelaPrecoDia
            </div>
        </section>
    </div>

    <script>
        function abrirAba(idAba, botao, nomeAba) {
            document.querySelectorAll('.tab-content').forEach(function(aba) {
                aba.classList.remove('active');
            });

            document.querySelectorAll('.tab-btn').forEach(function(btn) {
                btn.classList.remove('active');
            });

            document.getElementById(idAba).classList.add('active');
            botao.classList.add('active');

            if (nomeAba === 'Menor Preço') {
                document.getElementById('statAba').innerHTML = 'Menor Pre&ccedil;o';
            } else if (nomeAba === 'Planilha de Preços') {
                document.getElementById('statAba').innerHTML = 'Planilha de Pre&ccedil;os';
            } else {
                document.getElementById('statAba').innerHTML = 'Preco do dia';
            }

            atualizarStatsGerais();
        }

        function detectarIndices(tabelaId) {
            const ths = Array.from(document.querySelectorAll('#' + tabelaId + ' thead th'))
                .map(function(th) { return th.innerText.trim().toLowerCase(); });

            function acharIndiceExatoOuParcial(prioridades) {
                for (var i = 0; i < prioridades.length; i++) {
                    var termo = prioridades[i];
                    var exato = ths.findIndex(function(nome) { return nome === termo; });
                    if (exato >= 0) return exato;
                }

                for (var j = 0; j < prioridades.length; j++) {
                    var termo2 = prioridades[j];
                    var parcial = ths.findIndex(function(nome) { return nome.includes(termo2); });
                    if (parcial >= 0) return parcial;
                }

                return -1;
            }

            return {
                produto: acharIndiceExatoOuParcial(['produto', 'tipo', 'categoria']),
                modelo: acharIndiceExatoOuParcial(['modelo', 'nome modelo']),
                gb: acharIndiceExatoOuParcial(['gb', 'armazenamento', 'memoria', 'memória', 'capacidade']),
                condicao: acharIndiceExatoOuParcial(['condicao', 'condição', 'estado']),
                preco: acharIndiceExatoOuParcial(['preco', 'preço', 'valor', 'menorpreco', 'menor preço'])
            };
        }

        function preencherSelectComValores(selectId, valores) {
            var select = document.getElementById(selectId);
            if (!select) return;

            var valorAtual = select.value;
            var primeiroOption = select.options[0].outerHTML;
            select.innerHTML = primeiroOption;

            valores
                .filter(function(v) { return v && v.trim() !== ''; })
                .sort(function(a, b) {
                    return a.localeCompare(b, 'pt-BR', { numeric: true, sensitivity: 'base' });
                })
                .forEach(function(valor) {
                    var option = document.createElement('option');
                    option.value = valor;
                    option.textContent = valor;
                    select.appendChild(option);
                });

            select.value = valorAtual;
        }

        function extrairNumeroPreco(texto) {
            if (!texto) return NaN;
            var limpo = texto
                .toString()
                .replace(/R\$/gi, '')
                .replace(/\s+/g, '')
                .replace(/\./g, '')
                .replace(/,/g, '.')
                .replace(/[^\d.-]/g, '');
            return parseFloat(limpo);
        }

        function formatarPreco(valor) {
            if (isNaN(valor)) return '-';
            return 'R$ ' + valor.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        }

        function atualizarStatsGerais() {
            var abaAtiva = document.querySelector('.tab-content.active');
            if (!abaAtiva) return;

            var tabela = abaAtiva.querySelector('table');
            if (!tabela) return;

            var linhasVisiveis = Array.from(tabela.querySelectorAll('tbody tr'));
            document.getElementById('statRegistros').textContent = linhasVisiveis.length;

            var ths = Array.from(tabela.querySelectorAll('thead th')).map(function(th) {
                return th.innerText.trim().toLowerCase();
            });

            var idxModelo = ths.findIndex(function(t) {
                return t === 'modelo' || t.includes('modelo');
            });

            var idxPreco = ths.findIndex(function(t) {
                return t === 'preço' || t === 'preco' || t.includes('preço') || t.includes('preco') || t.includes('menorpreco');
            });

            var modelos = new Set();
            var precos = [];

            linhasVisiveis.forEach(function(linha) {
                var tds = linha.querySelectorAll('td');
                if (idxModelo >= 0 && tds[idxModelo]) modelos.add(tds[idxModelo].innerText.trim());
                if (idxPreco >= 0 && tds[idxPreco]) {
                    var n = extrairNumeroPreco(tds[idxPreco].innerText.trim());
                    if (!isNaN(n)) precos.push(n);
                }
            });

            document.getElementById('statModelos').textContent = modelos.size;
            document.getElementById('statMinPreco').textContent = precos.length ? formatarPreco(Math.min.apply(null, precos)) : '-';
            document.getElementById('statMaxPreco').textContent = precos.length ? formatarPreco(Math.max.apply(null, precos)) : '-';
        }

        function configurarFiltros(config) {
            var tabela = document.getElementById(config.tabelaId);
            if (!tabela) return;

            var tbody = tabela.querySelector('tbody');
            var busca = document.getElementById(config.buscaId);
            var produtoSelect = document.getElementById(config.produtoId);
            var modeloSelect = document.getElementById(config.modeloId);
            var gbSelect = document.getElementById(config.gbId);
            var condicaoSelect = document.getElementById(config.condicaoId);
            var precoMinInput = document.getElementById(config.precoMinId);
            var precoMaxInput = document.getElementById(config.precoMaxId);
            var ordenacaoSelect = document.getElementById(config.ordenacaoId);
            var resumo = document.getElementById(config.resumoId);
            var paginacao = document.getElementById(config.tabelaId + '_paginacao');

            var indices = detectarIndices(config.tabelaId);
            var jsonNode = document.getElementById(config.tabelaId + '_data');
            var dados = jsonNode ? JSON.parse(jsonNode.textContent) : [];

            var filtrados = dados.slice();
            var paginaAtual = 1;
            var pageSize = 50;
            var colunas = dados.length ? Object.keys(dados[0]) : [];

            function valorCampo(item, idx) {
                if (idx < 0 || idx >= colunas.length) return '';
                var chave = colunas[idx];
                var valor = item[chave];
                return (valor == null ? '' : valor).toString().trim();
            }

            function popularFiltros() {
                var produtos = new Set();
                var modelos = new Set();
                var gbs = new Set();
                var condicoes = new Set();

                dados.forEach(function(item) {
                    var produto = valorCampo(item, indices.produto);
                    var modelo = valorCampo(item, indices.modelo);
                    var gb = valorCampo(item, indices.gb);
                    var condicao = valorCampo(item, indices.condicao);

                    if (produto) produtos.add(produto);
                    if (modelo) modelos.add(modelo);
                    if (gb) gbs.add(gb);
                    if (condicao) condicoes.add(condicao);
                });

                preencherSelectComValores(config.produtoId, Array.from(produtos));
                preencherSelectComValores(config.modeloId, Array.from(modelos));
                preencherSelectComValores(config.gbId, Array.from(gbs));
                preencherSelectComValores(config.condicaoId, Array.from(condicoes));
            }

            function ordenarDados(lista) {
                var tipoOrdenacao = ordenacaoSelect ? ordenacaoSelect.value : '';
                if (!tipoOrdenacao || indices.preco < 0) return lista;

                return lista.slice().sort(function(a, b) {
                    var aPreco = extrairNumeroPreco(valorCampo(a, indices.preco));
                    var bPreco = extrairNumeroPreco(valorCampo(b, indices.preco));

                    var av = isNaN(aPreco) ? 0 : aPreco;
                    var bv = isNaN(bPreco) ? 0 : bPreco;

                    if (tipoOrdenacao === 'preco-asc') return av - bv;
                    if (tipoOrdenacao === 'preco-desc') return bv - av;
                    return 0;
                });
            }

            function renderTabela() {
                tbody.innerHTML = '';

                var inicio = (paginaAtual - 1) * pageSize;
                var fim = inicio + pageSize;
                var pagina = filtrados.slice(inicio, fim);

                var fragment = document.createDocumentFragment();

                pagina.forEach(function(item) {
                    var tr = document.createElement('tr');

                    colunas.forEach(function(col) {
                        var td = document.createElement('td');
                        td.textContent = (item[col] == null ? '' : item[col]).toString();
                        tr.appendChild(td);
                    });

                    fragment.appendChild(tr);
                });

                tbody.appendChild(fragment);
                renderPaginacao();
                atualizarStatsGerais();
            }

            function renderPaginacao() {
                if (!paginacao) return;

                var totalPaginas = Math.max(1, Math.ceil(filtrados.length / pageSize));

                var htmlPaginacao = '';
                htmlPaginacao += '<button ' + (paginaAtual <= 1 ? 'disabled' : '') + ' data-acao="prev">Anterior</button>';
                htmlPaginacao += '<span class="info-pagina">Página ' + paginaAtual + ' de ' + totalPaginas + '</span>';
                htmlPaginacao += '<button ' + (paginaAtual >= totalPaginas ? 'disabled' : '') + ' data-acao="next">Próxima</button>';

                paginacao.innerHTML = htmlPaginacao;

                var prev = paginacao.querySelector('[data-acao="prev"]');
                var next = paginacao.querySelector('[data-acao="next"]');

                if (prev) {
                    prev.onclick = function() {
                        if (paginaAtual > 1) {
                            paginaAtual--;
                            renderTabela();
                        }
                    };
                }

                if (next) {
                    next.onclick = function() {
                        if (paginaAtual < totalPaginas) {
                            paginaAtual++;
                            renderTabela();
                        }
                    };
                }
            }

            function aplicarFiltros() {
                var termo = (busca ? busca.value : '').toLowerCase().trim();
                var produto = (produtoSelect ? produtoSelect.value : '').toLowerCase().trim();
                var modelo = (modeloSelect ? modeloSelect.value : '').toLowerCase().trim();
                var gb = (gbSelect ? gbSelect.value : '').toLowerCase().trim();
                var condicao = (condicaoSelect ? condicaoSelect.value : '').toLowerCase().trim();
                var precoMin = parseFloat(precoMinInput ? precoMinInput.value : '');
                var precoMax = parseFloat(precoMaxInput ? precoMaxInput.value : '');

                filtrados = dados.filter(function(item) {
                    var texto = Object.values(item).join(' ').toLowerCase();

                    var valorProduto = valorCampo(item, indices.produto).toLowerCase();
                    var valorModelo = valorCampo(item, indices.modelo).toLowerCase();
                    var valorGb = valorCampo(item, indices.gb).toLowerCase();
                    var valorCondicao = valorCampo(item, indices.condicao).toLowerCase();
                    var valorPreco = extrairNumeroPreco(valorCampo(item, indices.preco));

                    var okBusca = !termo || texto.includes(termo);
                    var okProduto = !produto || valorProduto === produto;
                    var okModelo = !modelo || valorModelo === modelo;
                    var okGb = !gb || valorGb === gb;
                    var okCondicao = !condicao || valorCondicao === condicao;
                    var okPrecoMin = isNaN(precoMin) || (!isNaN(valorPreco) && valorPreco >= precoMin);
                    var okPrecoMax = isNaN(precoMax) || (!isNaN(valorPreco) && valorPreco <= precoMax);

                    return okBusca && okProduto && okModelo && okGb && okCondicao && okPrecoMin && okPrecoMax;
                });

                filtrados = ordenarDados(filtrados);
                paginaAtual = 1;

                if (resumo) {
                    resumo.textContent = filtrados.length + ' registro(s) encontrado(s)';
                }

                renderTabela();
            }

            var debounceTimer;
            function aplicarFiltrosComDebounce() {
                clearTimeout(debounceTimer);
                debounceTimer = setTimeout(aplicarFiltros, 180);
            }

            popularFiltros();
            aplicarFiltros();

            [busca, produtoSelect, modeloSelect, gbSelect, condicaoSelect, precoMinInput, precoMaxInput, ordenacaoSelect].forEach(function(el) {
                if (el) {
                    el.addEventListener('input', aplicarFiltrosComDebounce);
                    el.addEventListener('change', aplicarFiltros);
                }
            });
        }

        configurarFiltros({
            tabelaId: 'tabelaRelatorio',
            buscaId: 'buscaRelatorio',
            produtoId: 'produtoRelatorio',
            modeloId: 'modeloRelatorio',
            gbId: 'gbRelatorio',
            condicaoId: 'condicaoRelatorio',
            precoMinId: 'precoMinRelatorio',
            precoMaxId: 'precoMaxRelatorio',
            ordenacaoId: 'ordenacaoRelatorio',
            resumoId: 'resumoRelatorio'
        });

        configurarFiltros({
            tabelaId: 'tabelaPrecos',
            buscaId: 'buscaPrecos',
            produtoId: 'produtoPrecos',
            modeloId: 'modeloPrecos',
            gbId: 'gbPrecos',
            condicaoId: 'condicaoPrecos',
            precoMinId: 'precoMinPrecos',
            precoMaxId: 'precoMaxPrecos',
            ordenacaoId: 'ordenacaoPrecos',
            resumoId: 'resumoPrecos'
        });

        configurarFiltros({
            tabelaId: 'tabelaPrecoDia',
            buscaId: 'buscaPrecoDia',
            produtoId: 'produtoPrecoDia',
            modeloId: 'modeloPrecoDia',
            gbId: 'gbPrecoDia',
            condicaoId: 'condicaoPrecoDia',
            precoMinId: 'precoMinPrecoDia',
            precoMaxId: 'precoMaxPrecoDia',
            ordenacaoId: 'ordenacaoPrecoDia',
            resumoId: 'resumoPrecoDia'
        });

        atualizarStatsGerais();
    </script>
</body>
</html>
"@

$destino = Join-Path $docs "index.html"
[System.IO.File]::WriteAllText($destino, $html, [System.Text.UTF8Encoding]::new($false))

Write-Host "HTML gerado em docs\index.html - dashboard com aba Preco do dia"