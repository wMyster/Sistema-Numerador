SISTEMA UNICO DE NUMERADORES (OFICIOS, MEMORANDOS, ETC.)
===========================================================================
Gestao e Controle de Documentacao Oficial | Atualizado: 12/03/2026

Este sistema foi projetado para centralizar a emissao de numeros oficiais,
evitando erros de duplicidade e mantendo um historico digital perfeito.

1. LOGICA DE PERFORMANCE (SUPER RAPIDO!)
----------------------------------------
O sistema foi otimizado para lidar com grandes volumes de dados (5.000+):

- LIMITE DE TELA: Para manter a fluidez, a tabela carrega apenas os 200
  registros mais recentes. Isso evita travamentos ao rolar a lista.
- ORDEM DECRESCENTE: Os documentos novos aparecem sempre no topo da lista.
- BUSCA TOTAL: Mesmo mostrando apenas 200 registros na tela, a barra de 
  busca varre todos os milhares de registros no banco instantaneamente.

2. LOGICA DE NUMERACAO POR TIPO
-------------------------------
O sistema gerencia de forma independente a numeracao de 7 tipos:
- Oficios, Memorandos, Circulares, Notificacoes, Portarias, 
  Autorizacoes de Veiculo e Certidoes.

Cada tipo utiliza o banco de dados para sugerir o proximo numero oficial, 
garantindo que a sequencia nunca se perca ou se repita.

3. LIXEIRA E RECUPERACAO DE DADOS
---------------------------------
- SOFT-DELETE: Quando voce deleta um registro, ele vai para a Lixeira.
- RESTAURACAO: Voce pode recuperar um numero excluido por engano.
- EXCLUSAO DEFINITIVA: Limpa o banco definitivamente. Ao esvaziar a 
  lixeira, o sistema reseta os contadores para que os IDs continuem limpos.

4. AUDITORIA E DASHBOARD
------------------------
- REGISTRO DE ACOES: O sistema salva quem criou, quem editou e quando.
- PAINEL VISUAL: Veja estatisticas de quantos documentos foram gerados no
  mes atual e quem sao os funcionarios mais ativos.

5. INFRAESTRUTURA
-----------------
- BANCO DE DADOS: SQLite 3 com indices de alta velocidade.
- WORD INTEGRADO: Gera arquivos .docx preenchidos automaticamente.
- PORTABILIDADE: Funciona direto da pasta, sem precisar instalar Python.

COMO USAR:
- Abra pelo arquivo "run.bat" para uso normal.
- Use o "iniciar.vbs" para abrir o sistema de forma limpa e silenciosa.

---------------------------------------------------------------------------
Desenvolvido em colaboracao com Antigravity / Agentic Assistant.
Todo o sistema foi revisado para maxima estabilidade e performance.
