"""
Run solar financial model with formula-driven Excel output.
Creates interactive spreadsheet where all calculations happen in Excel.
"""

from model.inputs import load_inputs
from model.writer_formula_excel import write_formula_workbook


def run_formula_model(inputs_path: str = 'example_inputs.json',
                     output_path: str = 'SolarModel_Interactive.xlsx'):
    """
    Generate interactive formula-based Excel model.

    Args:
        inputs_path: Path to JSON inputs (used for initial values only)
        output_path: Path for output Excel file

    Returns:
        Path to generated Excel file
    """
    print("Generating interactive Solar Financial Model...")
    print(f"  Reading inputs from: {inputs_path}")

    # Load inputs (for initial values and validation)
    inputs, defaults_used, warnings = load_inputs(inputs_path)

    print(f"  Creating formula-based Excel workbook...")

    # Generate Excel workbook with formulas
    write_formula_workbook(inputs, defaults_used, warnings, output_path)

    print(f"  ✓ Excel workbook created: {output_path}")
    print()
    print("NEXT STEPS:")
    print(f"  1. Open {output_path}")
    print("  2. Go to 'Dashboard' tab")
    print("  3. Edit yellow input cells")
    print("  4. Watch green metrics update automatically!")

    return output_path


if __name__ == "__main__":
    run_formula_model()
