# ExceltoPDF - Conversor de Excel para PDF

## 🚀 Funcionalidades

### ✨ **Auto-geração de Nomes de Arquivo**
- **Automático**: O nome do arquivo PDF de saída é gerado automaticamente baseado no arquivo Excel de entrada
- **Prevenção de Sobrescrita**: Se um arquivo com o mesmo nome já existir, um número sequencial é adicionado (ex: `arquivo_1.pdf`, `arquivo_2.pdf`)
- **Tempo Real**: O nome é atualizado automaticamente sempre que você selecionar um novo arquivo de entrada

### 🔧 **Ajuste Automático de Dimensões**
- **AutoFit Inteligente**: Ajusta automaticamente a largura das colunas e altura das linhas baseado no conteúdo
- **Prevenção de Corte**: Habilita quebra de texto e ajusta dimensões para evitar corte de conteúdo
- **Orientação Otimizada**: Usa orientação paisagem para melhor aproveitamento do espaço
- **Margens Inteligentes**: Configura margens e espaçamentos para máxima legibilidade
- **Configurável**: Opção para habilitar/desabilitar o ajuste automático conforme sua necessidade
- **Ajuste Agressivo**: Modo extra que força dimensões maiores e mesclagem automática para casos difíceis (ATIVADO POR PADRÃO)

### 📁 **Como Funciona**
1. **Selecione o arquivo Excel**: Use o botão "Browse..." para escolher seu arquivo `.xlsx` ou `.xls`
2. **Nome automático**: O campo de saída será preenchido automaticamente com o mesmo nome, mas com extensão `.pdf`
3. **Personalização opcional**: Você pode alterar o nome manualmente se desejar
4. **Conversão**: Clique em "Convert" para iniciar a conversão

### 🔧 **Ajuste Automático de Dimensões - Como Funciona**
1. **Análise de Conteúdo**: O sistema analisa o conteúdo de cada coluna para determinar a largura ideal
2. **AutoFit Inteligente**: Aplica `AutoFit` nas colunas e linhas baseado no conteúdo real
3. **Quebra de Texto**: Habilita quebra de texto automática para células com conteúdo longo
4. **Orientação Paisagem**: Usa orientação paisagem para acomodar mais colunas na página
5. **Margens Otimizadas**: Configura margens e espaçamentos para máxima legibilidade
6. **Prevenção de Corte**: Garante que todo o conteúdo seja visível sem cortes

### 🔧 **Métodos de Conversão**
- **Auto**: Detecta automaticamente o melhor método disponível
- **Excel**: Usa o Microsoft Excel instalado (requer win32com)
- **ReportLab**: Usa bibliotecas Python (pandas + reportlab)

### 📋 **Opções Disponíveis**
- ✅ **Converter todas as abas**: Processa todas as planilhas do arquivo Excel
- ✅ **Saída verbosa**: Mostra informações detalhadas durante a conversão
- ✅ **Ajuste automático de dimensões**: Evita corte de texto ajustando automaticamente colunas e linhas (ATIVADO POR PADRÃO)
- ✅ **Ajuste agressivo**: Força dimensões maiores e mesclagem automática para casos difíceis (ATIVADO POR PADRÃO)
- ✅ **Interface gráfica responsiva**: Adapta-se automaticamente a diferentes tamanhos de janela
- ✅ **Interface gráfica intuitiva**: Fácil de usar para usuários não técnicos

## 🚀 **Como Executar**

### **Opção 1: Script Batch (Recomendado)**
```bash
.\run_exceltoPDF_gui.bat
```

### **Opção 2: Comando Direto**
```bash
exceltopdf-gui
```

### **Opção 3: Módulo Python**
```bash
python -m exceltopdf.gui
```

## 📥 **Instalação**

Se você ainda não instalou o projeto:
```bash
pip install -e .
```

## 💡 **Exemplos de Uso**

### **Exemplo 1: Conversão Simples**
- Arquivo de entrada: `relatorio_vendas.xlsx`
- Nome gerado automaticamente: `relatorio_vendas.pdf`

### **Exemplo 2: Prevenção de Sobrescrita**
- Arquivo de entrada: `relatorio_vendas.xlsx`
- Se `relatorio_vendas.pdf` já existir → `relatorio_vendas_1.pdf`
- Se `relatorio_vendas_1.pdf` já existir → `relatorio_vendas_2.pdf`

### **Exemplo 3: Múltiplas Abas**
- Marque "Converter todas as abas" para processar todas as planilhas do arquivo
- Cada aba será convertida para uma página separada no PDF

## 🛠️ **Requisitos do Sistema**

- Python 3.8 ou superior
- Bibliotecas: pandas, openpyxl, reportlab, click, PyPDF2
- Windows: win32com (opcional, para método Excel)

## 📞 **Suporte**

Se encontrar problemas:
1. Verifique se todas as dependências estão instaladas
2. Execute `python -m exceltopdf.gui --help` para ver opções disponíveis
3. Verifique os logs na interface para mensagens de erro detalhadas

---

**Desenvolvido com ❤️ para facilitar a conversão de arquivos Excel para PDF**
