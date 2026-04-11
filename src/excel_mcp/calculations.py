from typing import Any
import logging

from .workbook import require_worksheet, safe_workbook
from .cell_utils import validate_cell_reference
from .exceptions import ValidationError, CalculationError
from .validation import validate_formula

logger = logging.getLogger(__name__)

def apply_formula(
    filepath: str,
    sheet_name: str,
    cell: str,
    formula: str
) -> dict[str, Any]:
    """Apply any Excel formula to a cell."""
    try:
        if not validate_cell_reference(cell):
            raise ValidationError(f"Invalid cell reference: {cell}")

        # Ensure formula starts with =
        if not formula.startswith('='):
            formula = f'={formula}'

        # Validate formula syntax
        is_valid, message = validate_formula(formula)
        if not is_valid:
            raise CalculationError(f"Invalid formula syntax: {message}")

        with safe_workbook(filepath, save=True) as wb:
            sheet = require_worksheet(
                wb,
                sheet_name,
                error_cls=ValidationError,
                operation="applying formulas",
            )

            try:
                # Apply formula to the cell
                cell_obj = sheet[cell]
                cell_obj.value = formula
            except Exception as e:
                raise CalculationError(f"Failed to apply formula to cell: {str(e)}")

        return {
            "message": f"Applied formula '{formula}' to cell {cell}",
            "cell": cell,
            "formula": formula
        }

    except (ValidationError, CalculationError) as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to apply formula: {e}")
        raise CalculationError(str(e))
