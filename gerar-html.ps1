.getElementById(config.precoMinId);
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
                htmlPaginacao += '<button ' + (paginaAtual >= totalPaginas ? 'disabled' : '') + ' data-acao="next">Proxima</button>';

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

        configurarFiltros({
            tabelaId: 'tabelaPrecoOntem',
            buscaId: 'buscaPrecoOntem',
            produtoId: 'produtoPrecoOntem',
            modeloId: 'modeloPrecoOntem',
            gbId: 'gbPrecoOntem',
            condicaoId: 'condicaoPrecoOntem',
            precoMinId: 'precoMinPrecoOntem',
            precoMaxId: 'precoMaxPrecoOntem',
            ordenacaoId: 'ordenacaoPrecoOntem',
            resumoId: 'resumoPrecoOntem'
        });

        atualizarStatsGerais();
    </script>
</body>
</html>
"@

$destino = Join-Path $docs "index.html"
[System.IO.File]::WriteAllText($destino, $html, [System.Text.UTF8Encoding]::new($false))

Write-Host "HTML gerado em docs\index.html - dashboard com aba Preco do dia"