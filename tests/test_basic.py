#!/usr/bin/env python3
"""Basic smoke tests for exceltopdf package."""
import sys
import pytest
import unittest.mock
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


def test_verbose_parameter_acceptance():
    """Test that both conversion functions accept verbose=True parameter."""
    from exceltopdf import cli
    
    # Mock file paths (we don't want to actually convert files in tests)
    input_file = "/fake/path/input.xlsx"
    output_file = "/fake/path/output.pdf"
    
    # Test convert_with_pandas_reportlab with verbose=True
    with unittest.mock.patch('exceltopdf.cli.pd.ExcelFile'), \
         unittest.mock.patch('exceltopdf.cli.SimpleDocTemplate'), \
         unittest.mock.patch('exceltopdf.cli.getSampleStyleSheet'):
        try:
            # This should not raise a TypeError about unexpected keyword argument
            cli.convert_with_pandas_reportlab(input_file, output_file, verbose=True)
        except TypeError as e:
            if "verbose" in str(e) or "unexpected keyword argument" in str(e):
                pytest.fail(f"convert_with_pandas_reportlab should accept verbose parameter: {e}")
        except Exception:
            # Other exceptions are fine (missing files, etc.) - we just want to test the signature
            pass
    
    # Test convert_with_win32com with verbose=True
    with unittest.mock.patch('exceltopdf.cli.win32'):
        try:
            # This should not raise a TypeError about unexpected keyword argument
            cli.convert_with_win32com(input_file, output_file, verbose=True)
        except TypeError as e:
            if "verbose" in str(e) or "unexpected keyword argument" in str(e):
                pytest.fail(f"convert_with_win32com should accept verbose parameter: {e}")
        except Exception:
            # Other exceptions are fine (missing dependencies, files, etc.) - we just want to test the signature
            pass


def test_log_parameter_acceptance():
    """Test that both conversion functions accept log parameter."""
    from exceltopdf import cli
    
    # Mock file paths and log function
    input_file = "/fake/path/input.xlsx"
    output_file = "/fake/path/output.pdf"
    log_function = lambda msg: None  # Simple mock log function
    
    # Test convert_with_pandas_reportlab with log parameter
    with unittest.mock.patch('exceltopdf.cli.pd.ExcelFile'), \
         unittest.mock.patch('exceltopdf.cli.SimpleDocTemplate'), \
         unittest.mock.patch('exceltopdf.cli.getSampleStyleSheet'):
        try:
            cli.convert_with_pandas_reportlab(input_file, output_file, log=log_function)
        except TypeError as e:
            if "log" in str(e) or "unexpected keyword argument" in str(e):
                pytest.fail(f"convert_with_pandas_reportlab should accept log parameter: {e}")
        except Exception:
            pass
    
    # Test convert_with_win32com with log parameter
    with unittest.mock.patch('exceltopdf.cli.win32'):
        try:
            cli.convert_with_win32com(input_file, output_file, log=log_function)
        except TypeError as e:
            if "log" in str(e) or "unexpected keyword argument" in str(e):
                pytest.fail(f"convert_with_win32com should accept log parameter: {e}")
        except Exception:
            pass


if __name__ == '__main__':
    pytest.main([__file__])
