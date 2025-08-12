# ExceltoPDF

[🇧🇷 Português](README.md) | [🇺🇸 English](README_EN.md)

Uma ferramenta com interface gráfica e linha de comando para converter arquivos Excel para PDF com formatação otimizada, garantindo que todas as colunas caibam em uma página por planilha.

> **⚠️ Aviso Importante:** Este pacote ainda não está disponível no PyPI. Para instalar, clone o repositório e instale a partir do código fonte.

## Funcionalidades

• **Ajuste Inteligente de Colunas**: Ajusta automaticamente todas as colunas para caber na largura de uma página por planilha
• **Múltiplos Métodos de Conversão**:
  • Windows com Excel instalado: Usa win32com para exportação nativa do Excel para PDF
  • Alternativa multiplataforma: Usa pandas + reportlab para compatibilidade universal
• **Processamento em Lote**: Processa múltiplas planilhas em um único arquivo Excel
• **Interface de Linha de Comando**: Fácil de usar no terminal ou scripts
• **Interface Gráfica**: GUI amigável com seletores de arquivo e opções
• **Saída Flexível**: Mantém a integridade dos dados enquanto otimiza o layout

## Instalação

### A Partir do Código Fonte (recomendado)

```bash
git clone https://github.com/dadebr/ExceltoPDF.git
cd ExceltoPDF
pip install -e .
```

### Dependências

A ferramenta usará automaticamente o melhor método disponível:

Para Windows com Microsoft Excel:
```bash
pip install pywin32
```

Para compatibilidade multiplataforma:
```bash
pip install pandas openpyxl reportlab
```

Todas as dependências estão listadas em requirements.txt e serão instaladas automaticamente.

## Uso

### Interface Gráfica

Para executar a interface gráfica:
```bash
exceltopdf-gui
```

Alternativamente, se o comando não estiver disponível:
```bash
python -m exceltopdf.gui
```

#### Funcionalidades da Interface Gráfica

• **Seleção de Arquivos**: Botões de navegação para escolher arquivos Excel de entrada e PDF de saída
• **Métodos de Conversão**: Menu suspenso com opções:
  • auto - Detecta automaticamente o melhor método
  • excel - Usa o Excel nativo (Windows)
  • reportlab - Usa pandas + reportlab (multiplataforma)
• **Opções de Saída**: Caixa de seleção para habilitar saída detalhada
• **Converter Todas as Abas**: Caixa de seleção "Converter todas as abas" para processar todas as planilhas em um único PDF
• **Área de Log**: Mostra o progresso da conversão e detalhes em tempo real
• **Barra de Progresso**: Indicador visual durante o processo de conversão

#### Como Usar a Interface Gráfica

1. Execute `exceltopdf-gui` no terminal
2. Clique em "Browse..." ao lado de "Input Excel File" para selecionar seu arquivo Excel
3. Clique em "Browse..." ao lado de "Output PDF File" para escolher onde salvar o PDF
4. Selecione o método de conversão desejado no menu suspenso
5. Marque "Verbose output" se quiser informações detalhadas
6. Marque "Converter todas as abas" se quiser processar todas as planilhas
7. Clique em "Convert" para iniciar a conversão
8. Monitore o progresso na área de log

A interface executa em uma thread separada para prevenir travamento durante a conversão e exibe mensagens de sucesso ou erro ao final do processo.

### Interface de Linha de Comando

#### Uso Básico

```bash
# Converter arquivo Excel para PDF
exceltopdf input.xlsx output.pdf

# Com saída detalhada
exceltopdf input.xlsx output.pdf --verbose
```

#### Opções Avançadas

```bash
# Forçar método de conversão específico
exceltopdf input.xlsx output.pdf --method win32com
exceltopdf input.xlsx output.pdf --method pandas

# Detectar automaticamente o melhor método (padrão)
exceltopdf input.xlsx output.pdf --method auto

# Converter todas as abas do arquivo Excel para um único PDF
exceltopdf input.xlsx output.pdf --all-sheets

# Combinar opções
exceltopdf input.xlsx output.pdf --all-sheets --verbose --method auto
```

### API Python

```python
from exceltopdf.cli import convert_with_pandas_reportlab, convert_with_win32com

# Usando pandas/reportlab (multiplataforma)
convert_with_pandas_reportlab('input.xlsx', 'output.pdf')

# Usando win32com (Windows + Excel apenas)
convert_with_win32com('input.xlsx', 'output.pdf')
```

## Formatos Suportados

• **Entrada**: .xlsx, .xls
• **Saída**: .pdf

## Funcionamento

### Método 1: win32com (Windows + Excel)

• Usa a funcionalidade de exportação PDF integrada do Microsoft Excel
• Configura a configuração da página para ajustar todas as colunas em uma página
• Fornece a saída de maior qualidade com formatação nativa
• Aplica automaticamente escalonamento para garantir que as colunas se ajustem

### Método 2: pandas + reportlab (Multiplataforma)

• Lê dados do Excel usando pandas
• Converte para PDF usando reportlab
• Calcula automaticamente as larguras das colunas para caber na página
• Funciona em qualquer plataforma sem o Excel instalado

## Exemplos

```bash
# Conversão simples
exceltopdf sales_report.xlsx sales_report.pdf

# Conversão com log detalhado
exceltopdf financial_data.xlsx financial_data.pdf -v

# Forçar método multiplataforma
exceltopdf data.xlsx output.pdf --method pandas

# Converter todas as abas para um único PDF
exceltopdf workbook.xlsx complete_report.pdf --all-sheets
```

## Desenvolvimento

### Configurar Ambiente de Desenvolvimento

```bash
git clone https://github.com/dadebr/ExceltoPDF.git
cd ExceltoPDF
pip install -e .[dev]
```

### Executar Testes

```bash
# Executar todos os testes
pytest

# Executar com cobertura
pytest --cov=exceltopdf

# Executar arquivo de teste específico
pytest tests/test_basic.py
```

### Build do Pacote

```bash
# Construir pacotes de distribuição
python -m build

# Upload para PyPI (apenas mantenedores)
twine upload dist/*
```

## Contribuição

1. Faça um fork do repositório
2. Crie uma branch para sua funcionalidade (`git checkout -b feature/funcionalidade-incrivel`)
3. Faça commit das suas mudanças (`git commit -m 'Adicionar funcionalidade incrível'`)
4. Faça push para a branch (`git push origin feature/funcionalidade-incrivel`)
5. Abra um Pull Request

## Licença

Este projeto está licenciado sob a Licença MIT - veja o arquivo [LICENSE](LICENSE) para detalhes.

## Changelog

### v0.1.0

• Lançamento inicial
• Conversão básica de Excel para PDF
• Compatibilidade multiplataforma
• Interface de linha de comando
• Ajuste automático de colunas

## Solução de Problemas

### Problemas Comuns

**"Failed to import win32com"**
• Instale pywin32: `pip install pywin32`
• Ou use o método pandas: `--method pandas`

**"Required packages not available"**
• Instale as dependências: `pip install pandas openpyxl reportlab`

**"Input file does not exist"**
• Verifique o caminho do arquivo e certifique-se de que o arquivo existe
• Use caminhos absolutos se necessário

**Saída PDF está cortada**
• A ferramenta ajusta automaticamente as colunas, mas planilhas muito largas podem precisar de ajuste manual
• Considere usar orientação paisagem no arquivo Excel de origem

## Suporte

Se você encontrar problemas ou tiver dúvidas:

1. Verifique a [seção de solução de problemas](#solução-de-problemas)
2. Pesquise [issues existentes](https://github.com/dadebr/ExceltoPDF/issues)
3. Crie uma [nova issue](https://github.com/dadebr/ExceltoPDF/issues/new) com detalhes sobre seu problema

Feito com ❤️ para facilitar a conversão de Excel para PDF
