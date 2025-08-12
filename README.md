# ExceltoPDF

A command-line tool to convert Excel files to PDF with optimized formatting that ensures all columns fit on one page per worksheet.

## Features

• **Smart Column Fitting**: Automatically fits all columns to one page width per worksheet
• **Multiple Conversion Methods**: 
  • Windows with Excel installed: Uses win32com for native Excel PDF export
  • Cross-platform fallback: Uses pandas + reportlab for universal compatibility
• **Batch Processing**: Process multiple worksheets in a single Excel file
• **Command-line Interface**: Easy to use from terminal or scripts
• **Graphical Interface**: User-friendly GUI with file pickers and options
• **Flexible Output**: Maintains data integrity while optimizing layout

## Installation

### From PyPI (recommended)

```bash
pip install exceltopdf
```

### From Source

```bash
git clone https://github.com/dadebr/ExceltoPDF.git
cd ExceltoPDF
pip install -e .
```

### Dependencies

The tool will automatically use the best available method:

**For Windows with Microsoft Excel:**
```bash
pip install pywin32
```

**For cross-platform compatibility:**
```bash
pip install pandas openpyxl reportlab
```

All dependencies are listed in `requirements.txt` and will be installed automatically.

## Usage

### Basic Usage

```bash
# Convert Excel file to PDF
exceltopdf input.xlsx output.pdf

# With verbose output
exceltopdf input.xlsx output.pdf --verbose
```

### Advanced Options

```bash
# Force specific conversion method
exceltopdf input.xlsx output.pdf --method win32com
exceltopdf input.xlsx output.pdf --method pandas

# Auto-detect best method (default)
exceltopdf input.xlsx output.pdf --method auto
```

### Python API

```python
from exceltopdf.cli import convert_with_pandas_reportlab, convert_with_win32com

# Using pandas/reportlab (cross-platform)
convert_with_pandas_reportlab('input.xlsx', 'output.pdf')

# Using win32com (Windows + Excel only)
convert_with_win32com('input.xlsx', 'output.pdf')
```

## Interface Gráfica

Esta ferramenta oferece uma interface gráfica amigável baseada em Tkinter para usuários que preferem uma experiência visual.

### Instalação da Interface Gráfica

A interface gráfica é incluída automaticamente quando você instala o pacote:

```bash
pip install exceltopdf
```

### Uso da Interface Gráfica

Para iniciar a interface gráfica, execute o comando:

```bash
exceltopdf-gui
```

### Recursos da Interface Gráfica

• **Seleção de Arquivos**: Botões "Browse" para escolher arquivos Excel de entrada e PDF de saída
• **Métodos de Conversão**: Menu dropdown com opções:
  • `auto` - Detecta automaticamente o melhor método
  • `excel` - Usa Excel nativo (Windows)
  • `reportlab` - Usa pandas + reportlab (multiplataforma)
• **Opções de Saída**: Checkbox para habilitar saída verbose
• **Área de Log**: Mostra o progresso e detalhes da conversão em tempo real
• **Barra de Progresso**: Indicador visual durante o processo de conversão

### Como Usar a Interface Gráfica

1. Execute `exceltopdf-gui` no terminal
2. Clique em "Browse..." ao lado de "Input Excel File" para selecionar seu arquivo Excel
3. Clique em "Browse..." ao lado de "Output PDF File" para escolher onde salvar o PDF
4. Selecione o método de conversão desejado no dropdown
5. Marque "Verbose output" se desejar ver informações detalhadas
6. Clique em "Convert" para iniciar a conversão
7. Acompanhe o progresso na área de log

A interface roda em thread separada para não travar durante a conversão e exibe mensagens de sucesso ou erro ao final do processo.

## Supported File Formats

• Input: .xlsx, .xls
• Output: .pdf

## How It Works

### Method 1: win32com (Windows + Excel)

• Uses Microsoft Excel's built-in PDF export functionality
• Configures page setup to fit all columns on one page
• Provides the highest quality output with native formatting
• Automatically applies scaling to ensure columns fit

### Method 2: pandas + reportlab (Cross-platform)

• Reads Excel data using pandas
• Converts to PDF using reportlab
• Automatically calculates column widths to fit page
• Works on any platform without Excel installed

## Examples

```bash
# Simple conversion
exceltopdf sales_report.xlsx sales_report.pdf

# Convert with verbose logging
exceltopdf financial_data.xlsx financial_data.pdf -v

# Force cross-platform method
exceltopdf data.xlsx output.pdf --method pandas
```

## Development

### Setup Development Environment

```bash
git clone https://github.com/dadebr/ExceltoPDF.git
cd ExceltoPDF
pip install -e .[dev]
```

### Running Tests

```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=exceltopdf

# Run specific test file
pytest tests/test_basic.py
```

### Building Package

```bash
# Build distribution packages
python -m build

# Upload to PyPI (maintainers only)
twine upload dist/*
```

## Contributing

1. Fork the repository
2. Create a feature branch (git checkout -b feature/amazing-feature)
3. Commit your changes (git commit -m 'Add some amazing feature')
4. Push to the branch (git push origin feature/amazing-feature)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Changelog

### v0.1.0

• Initial release
• Basic Excel to PDF conversion
• Cross-platform compatibility
• Command-line interface
• Automatic column fitting

## Troubleshooting

### Common Issues

**"Failed to import win32com"**
• Install pywin32: pip install pywin32
• Or use pandas method: --method pandas

**"Required packages not available"**
• Install dependencies: pip install pandas openpyxl reportlab

**"Input file does not exist"**
• Check file path and ensure file exists
• Use absolute paths if needed

**PDF output is cut off**
• The tool automatically fits columns, but very wide spreadsheets may need manual adjustment
• Consider using landscape orientation in source Excel file

## Support

If you encounter any issues or have questions:

1. Check the [troubleshooting section](#troubleshooting)
2. Search [existing issues](https://github.com/dadebr/ExceltoPDF/issues)
3. Create a [new issue](https://github.com/dadebr/ExceltoPDF/issues/new) with details about your problem

Made with ❤️ for easier Excel to PDF conversion
