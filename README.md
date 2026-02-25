# 🏛️ Sistema Único de Numeradores

Um sistema completo de gestão de numeração e emissão de documentos oficiais desenvolvido em Python com Tkinter (Interface Gráfica) e SQLite (Banco de Dados).
Criado para otimizar o fluxo de trabalho de departamentos públicos ou privados, garantindo a integridade, controle e histórico de todos os Ofícios, Memorandos, Circulares, Notificações, Portarias, Autorizações de Veículo e Certidões emitidos pelo Setor.

![Python](https://img.shields.io/badge/python-3.10+-blue.svg)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-lightgrey.svg)
![SQLite](https://img.shields.io/badge/Database-SQLite3-blue.svg)

---

## 🚀 Funcionalidades Principais

- **Múltiplos Documentos**: Gerencia nativamente 7 tipos de documentos oficiais em abas independentes e expansíveis.
- **Auto-Incremento Inteligente**: Identifica o último número utilizado no banco (mesmo após exclusões parciais) e sugere automaticamente a próxima numeração oficial, evitando pular valores ou sobrepor registros.
- **Controle de Acesso Descomplicado**: Tela de login super amigável com Lista Expansível (Listbox) de funcionários do setor. Suporte para adição ou exclusão de perfis (proteção contra deleção de contas nativas como TI).
- **Auditoria e Logs em Tempo Real**: O sistema mantém um histórico completo e irremovível de quem criou, editou ou apagou qualquer documento registrado, marcando o horário exato da alteração (Aba interna secreta `REGISTRO`).
- **Geração de Documentos `.docx` Automáticos**: Emite documentos baseados em gabaritos limpos (templates) no Office, preenchendo as Tags com os metadados exatos do painel selecionado de um número. (Aba `RELATÓRIOS`).
- **Pesquisa Instantânea Inteligente e Responsiva**: Busca poderosa em todas as colunas visíveis do painel a cada letra digitada, permitindo achar placas de carros, assuntos ou destinatários antigas em milissegundos.
- **Auto-Backup e Sincronização em Rede**: 
  - Banco de Dados Central preparado para atuar na nuvem de Redes Compartilhadas da Prefeitura com inteligência *Fallback* (Modo Offline Secundário se a rede cair).
  - Rotinas silenciosas de Cópia de Segurança ativadas **a cada 15 minutos** de inatividade. O sistema disparará de forma invisível clones de tabelas preservando dezenas de Históricos Locais com carimbos para reverter perdas.
  - Sincroniza dinamicamente computadores em tempo real para manter os N°s Oficiais unificados quando várias pessoas acessam simultaneamente.
- **Imunidade a Telas "Fantasmas" e Bugs do SO**: Focos interativos nativos do Sistema Operacional travados na raiz para evitar comportamentos inesperados (Campos de formulários sendo selecionados acidentalmente via TAB, cintilações nas transições).

---

## 🧰 Tecnologias Utilizadas

- **Linguagem Principal**: `Python`
- **Interface Gráfica**: `Tkinter` (Customizado com CSS/Temas nativos)
- **Banco de Dados**: `SQLite 3`
- **Operações Office**: Integração profunda das APIs com a biblioteca externa `python-docx` para edição de templates do Word nativamente do app.
- **Shell Scripting & Automação**: `.vbs` e `.bat` integrados para execução portátil silenciosas sem forçar consoles do prompt CMD nas telas.

---

## 📁 Estrutura de Diretórios 

```text
Sistema-Numerador/
├── app/
│   ├── main.py          # Script Master que orquestra e estipula regras de carregamento
│   ├── ui.py            # Todos os 10+ Componentes das Interfaces, Abas e Threads 
│   ├── db.py            # Lógica das Consultas SQL, Migração Externa e Funções de Auditoria
│   └── export_docx.py   # Motor e Parser para Injeção de Tags dentro dos Templates Word
├── run.bat              # Atalho simplificado de console nativo
├── iniciar.vbs          # Wrapper Executável mudo (Sem console/Background mode)
├── .gitignore           # Ignorar ambientes/compiladores pesados nativos
└── README.md            # Documentação Central
```
> *Pastas ausentes no repositório (Criadas automaticamente em Runtime)*: `data/` (Bancos locais), `backup/` (Cópias de Segurança), e a `Numeradores/` (Para que você suba localmente o arquivo de Template que sua prefeitura ou corporação utilize lá dentro dos gabaritos da aplicação). 

---

## ⚙️ Instalação / Uso para Desenvolvedores

O aplicativo original foi arquitetado no modelo **Portable**, sem exigir que a pessoa final instalasse as variáveis de Python na máquina Windows de produção. Ele carregava todo o interpretador contido (`runtime/`) na aba principal (Que foi ignorado via `.gitignore` aqui no site por pesar em torno de 100 Megabytes e não haver limite para ele).

*Caso baixe o projeto cru (apenas o repositório fonte desta Master para estudar)*:

1. Certifique-se de que a sua máquina / SO possui os paths do [Python 3.10+](https://www.python.org/downloads/) configurados.
2. É recomendável criar uma *Virtual Environment* (VENV) no terminal da sua IDE para evitar poluição do Windows local.
3. Instale localmente o motor de templates do Office necessário:
   ```bash
   pip install python-docx
   ```
4. Navegue no terminal até a pasta raiz e rode o programa principal através do `app` diretamente:
   ```bash
   python app/main.py
   ```
5. *(Opcional)*: Use os atalhos de inicialização root `.bat` ou o `.vbs` inclusos para uma execução invisível em ambiente de Produção e Uso da Secretaria.

*Disclaimer:* O Painel do BD base inicial (`data/numerador.sqlite`) será gerado automaticamente do absoluto zero no HD assim que o comando supracitado engatilhar pela primeira vez, reconstruindo os layouts.

---
**Desenvolvido em colaboração através de Pair Programming com Antigravity / Agentic Assistant.** - Engenharias avançadas de interface ao vivo.
