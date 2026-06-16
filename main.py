"""
main.py - DocSkill CLI entry point.

Usage:
  Interactive:   python main.py
  CLI:           python main.py -i input_docs -o output_docs --generate-figures
"""

import argparse
import os
import sys
from doc_parser import DocumentParser
from formatter import DocumentFormatter
from figure_generator import FigureGenerator


SUPPORTED_EXTENSIONS = {'.docx', '.md', '.markdown', '.txt', '.html', '.htm', '.tex', '.rst', '.epub', '.odt'}


def process_file(input_path, output_path, generate_figures):
    """Process a single document file."""
    temp_docx = os.path.join(os.path.dirname(output_path) or '.', '_temp_processing.docx')

    try:
        print(f"\n{'='*50}")
        print(f"Processing: {os.path.basename(input_path)}")
        print(f"{'='*50}")

        # Step 1: Parse / convert to DOCX
        print("\n[1/3] Parsing document...")
        parser = DocumentParser(input_path)
        parser.parse_to_docx(temp_docx)

        # Step 2: (Optional) Generate figures from tables
        if generate_figures:
            print("\n[2/3] Generating figures from tables...")
            fig_gen = FigureGenerator(temp_docx)
            fig_gen.generate_and_insert_figures(output_dir='output_figures')
        else:
            print("\n[2/3] Skipping figure generation (use --generate-figures to enable)")

        # Step 3: Apply professional formatting
        print("\n[3/3] Applying professional formatting...")
        formatter = DocumentFormatter(temp_docx)
        formatter.format(output_path)

        print(f"\n[OK] Done! Output: {output_path}")

    except Exception as e:
        print(f"\n[ERROR] {e}")
        import traceback
        traceback.print_exc()

    finally:
        if os.path.exists(temp_docx):
            os.remove(temp_docx)


def interactive_mode():
    """Interactive menu when run without arguments."""
    print()
    print("=" * 50)
    print("       DocSkill - Document Formatter")
    print("=" * 50)
    print()

    # --- Input ---
    choice = input("1. Process all files in 'input_docs/' folder? [Y/n]: ").strip().lower()
    if choice == 'n':
        input_path = input("   Enter path to file or directory: ").strip().strip('"')
    else:
        input_path = "input_docs"

    # --- Output ---
    choice = input("2. Save outputs to 'output_docs/' folder? [Y/n]: ").strip().lower()
    if choice == 'n':
        output_path = input("   Enter output folder path: ").strip().strip('"')
    else:
        output_path = "output_docs"

    # --- Figures ---
    choice = input("3. Auto-generate plots from numeric tables? [y/N]: ").strip().lower()
    generate_figures = choice == 'y'

    print()
    return input_path, output_path, generate_figures


def run(input_path, output_path, generate_figures):
    """Core dispatcher: handles single file or directory."""
    if not os.path.exists(input_path):
        print(f"Error: '{input_path}' not found.")
        print("Place your documents in the 'input_docs/' folder and try again.")
        sys.exit(1)

    # --- Single file mode ---
    if os.path.isfile(input_path):
        # If output looks like a directory (no extension), put file inside it
        _, ext = os.path.splitext(output_path)
        if not ext:
            os.makedirs(output_path, exist_ok=True)
            name = os.path.splitext(os.path.basename(input_path))[0]
            out_file = os.path.join(output_path, f"{name}_formatted.docx")
        else:
            out_dir = os.path.dirname(output_path)
            if out_dir:
                os.makedirs(out_dir, exist_ok=True)
            out_file = output_path

        process_file(input_path, out_file, generate_figures)
        return

    # --- Directory mode ---
    if os.path.isdir(input_path):
        os.makedirs(output_path, exist_ok=True)

        files = sorted([
            f for f in os.listdir(input_path)
            if os.path.isfile(os.path.join(input_path, f))
            and not f.startswith('~$')
            and not f.startswith('.')
            and os.path.splitext(f)[1].lower() in SUPPORTED_EXTENSIONS
        ])

        if not files:
            print(f"No supported documents found in '{input_path}/'.")
            print(f"Supported formats: {', '.join(sorted(SUPPORTED_EXTENSIONS))}")
            return

        print(f"\nFound {len(files)} document(s) to process.")

        for f in files:
            in_file = os.path.join(input_path, f)
            name = os.path.splitext(f)[0]
            out_file = os.path.join(output_path, f"{name}_formatted.docx")
            process_file(in_file, out_file, generate_figures)

        print(f"\n{'='*50}")
        print(f"All done! Check '{output_path}/' for results.")
        print(f"{'='*50}")


def main():
    if len(sys.argv) == 1:
        # No arguments → interactive mode
        in_path, out_path, gen_figs = interactive_mode()
        run(in_path, out_path, gen_figs)
    else:
        ap = argparse.ArgumentParser(
            description="DocSkill - Professional Document Formatter & Figure Generator"
        )
        ap.add_argument("-i", "--input", default="input_docs",
                        help="Input file or directory (default: input_docs/)")
        ap.add_argument("-o", "--output", default="output_docs",
                        help="Output file or directory (default: output_docs/)")
        ap.add_argument("--generate-figures", action="store_true",
                        help="Auto-generate plots from numeric tables")
        args = ap.parse_args()
        run(args.input, args.output, args.generate_figures)


if __name__ == "__main__":
    main()
