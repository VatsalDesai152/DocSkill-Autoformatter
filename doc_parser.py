"""
doc_parser.py - Document parsing and conversion module.

Converts various document formats to DOCX using Pandoc,
with proper math rendering support.
"""

import os
import shutil


class DocumentParser:
    """Parses input documents and converts them to DOCX format."""

    # Formats that pypandoc can handle
    SUPPORTED_EXTENSIONS = ['.md', '.markdown', '.txt', '.html', '.htm', '.tex', '.rst', '.epub', '.odt']

    def __init__(self, input_path):
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
        self.input_path = os.path.abspath(input_path)

    def parse_to_docx(self, output_docx_path):
        """
        Convert the input file to a .docx intermediate.
        For .docx inputs, just copies the file.
        For everything else, uses Pandoc via pypandoc.
        """
        ext = os.path.splitext(self.input_path)[1].lower()

        if ext == '.docx':
            shutil.copy2(self.input_path, output_docx_path)
            print(f"  Loaded DOCX: {os.path.basename(self.input_path)}")
            return output_docx_path

        if ext not in self.SUPPORTED_EXTENSIONS:
            print(f"  Warning: '{ext}' may not be fully supported. Attempting anyway...")

        # Ensure pypandoc and pandoc are available
        pypandoc = self._get_pypandoc()

        # Build conversion arguments
        input_format = self._get_input_format(ext)
        extra_args = []

        print(f"  Converting {ext} -> DOCX via Pandoc (format: {input_format or 'auto'})...")

        pypandoc.convert_file(
            self.input_path,
            'docx',
            format=input_format,
            outputfile=output_docx_path,
            extra_args=extra_args,
        )

        return output_docx_path

    @staticmethod
    def _get_pypandoc():
        """Import pypandoc and ensure the pandoc binary exists."""
        try:
            import pypandoc
        except ImportError:
            raise ImportError(
                "pypandoc is required. Install it with: pip install pypandoc"
            )

        # Check if pandoc binary is available; download if missing
        try:
            version = pypandoc.get_pandoc_version()
            print(f"  Using Pandoc {version}")
        except OSError:
            print("  Pandoc binary not found. Downloading automatically...")
            pypandoc.download_pandoc()
            print(f"  Downloaded Pandoc {pypandoc.get_pandoc_version()}")

        return pypandoc

    @staticmethod
    def _get_input_format(ext):
        """
        Return the Pandoc input format string with math extensions enabled.
        Pandoc converts LaTeX math ($...$, $$...$$) to native Word OMML
        equations by default — no special flags needed beyond enabling
        the correct math extensions in the input format.
        """
        if ext in ['.md', '.markdown']:
            # Enable all math syntaxes so Pandoc recognises them
            return (
                'markdown'
                '+tex_math_dollars'
                '+tex_math_single_backslash'
                '+tex_math_double_backslash'
            )
        if ext == '.tex':
            return 'latex'
        if ext in ['.html', '.htm']:
            return 'html'
        if ext == '.rst':
            return 'rst'
        # For everything else, let Pandoc auto-detect
        return None
