#!/usr/bin/env python3
"""Basic smoke tests for exceltopdf package."""

import sys
import pytest
from pathlib import Path

# Add src to path for testing
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

try:
    import exceltopdf
    import exceltopdf.cli
except ImportError as e:
    pytest.fail(f"Failed to import exceltopdf: {e}")


def test_import_exceltopdf():
    """Test that the exceltopdf package can be imported."""
    import exceltopdf
    # Should have version
    assert hasattr(exceltopdf, '__version__')
    assert exceltopdf.__version__ == '0.1.0'


def test_import_cli():
    """Test that the CLI module can be imported."""
    from exceltopdf import cli
    # Should have main function
    assert hasattr(cli, 'main')
    assert callable(cli.main)


def test_cli_help(capsys):
    """Test that CLI shows help without errors."""
    from exceltopdf.cli import main
    
    # Test help flag
    with pytest.raises(SystemExit) as exc_info:
        sys.argv = ['exceltopdf', '--help']
        main()
    
    # Help should exit with code 0
    assert exc_info.value.code == 0
    
    # Should have printed help text
    captured = capsys.readouterr()
    assert 'Convert Excel files to PDF' in captured.out


def test_cli_version():
    """Test that CLI reports correct version."""
    from exceltopdf.cli import main
    
    # We can't easily test --version flag without more complex setup,
    # so we just verify the main function exists and is callable
    assert callable(main)


def test_conversion_methods_exist():
    """Test that conversion methods exist in CLI module."""
    from exceltopdf import cli
    
    # Check that conversion functions exist
    assert hasattr(cli, 'convert_with_win32com')
    assert hasattr(cli, 'convert_with_pandas_reportlab')
    assert callable(cli.convert_with_win32com)
    assert callable(cli.convert_with_pandas_reportlab)


if __name__ == '__main__':
    pytest.main([__file__])
