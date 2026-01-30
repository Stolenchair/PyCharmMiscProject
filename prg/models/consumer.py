"""Consumer data model."""

from dataclasses import dataclass
from typing import Optional


@dataclass
class ConsumerData:
    """
    Consumer data structure.

    Represents a gas consumer (either population or organization) with their
    expenses and bindings to PRG pipelines.
    """
    # Identity
    id: str
    type: str  # "Население" (population) or "Организация" (organization)
    name: str

    # Location
    mo: str  # Municipal district
    settlement: str  # Settlement

    # Binding and expenses
    code: str  # PRG binding string: "PRG_ID:share:GRS_name;PRG_ID2:share2:GRS_name2"
    expenses: Optional[float] = None  # Yearly expenses
    hourly_expenses: Optional[float] = None  # Hourly expenses

    # Organization-specific
    grs_id: Optional[str] = None  # For organizations only
    grs_name: Optional[str] = None  # For organizations only

    # Excel metadata (for persistence)
    sheet_name: str = ""
    excel_row: int = 0
    name_col: Optional[int] = None  # Only for population
    code_col: int = 0
    expenses_col: int = 0
    hourly_expenses_col: Optional[int] = None
    grs_id_col: Optional[int] = None  # Only for organizations

    def __repr__(self):
        return f"ConsumerData(type='{self.type}', name='{self.name}', mo='{self.mo}')"

    def has_expenses(self) -> bool:
        """Check if consumer has any expense data."""
        return (self.expenses is not None and self.expenses > 0) or \
               (self.hourly_expenses is not None and self.hourly_expenses > 0)

    def is_population(self) -> bool:
        """Check if consumer is population type."""
        return self.type == "Население"

    def is_organization(self) -> bool:
        """Check if consumer is organization type."""
        return self.type == "Организация"

    def to_dict(self):
        """Convert to dictionary (for compatibility with existing code)."""
        result = {
            'id': self.id,
            'type': self.type,
            'name': self.name,
            'mo': self.mo,
            'settlement': self.settlement,
            'code': self.code,
            'expenses': self.expenses,
            'hourly_expenses': self.hourly_expenses,
            'sheet_name': self.sheet_name,
            'excel_row': self.excel_row,
            'code_col': self.code_col,
            'expenses_col': self.expenses_col,
        }

        if self.name_col is not None:
            result['name_col'] = self.name_col
        if self.hourly_expenses_col is not None:
            result['hourly_expenses_col'] = self.hourly_expenses_col
        if self.grs_id is not None:
            result['grs_id'] = self.grs_id
        if self.grs_name is not None:
            result['grs_name'] = self.grs_name
        if self.grs_id_col is not None:
            result['grs_id_col'] = self.grs_id_col

        return result

    @classmethod
    def from_dict(cls, data: dict):
        """Create from dictionary (for compatibility with existing code)."""
        # Handle optional fields
        kwargs = data.copy()
        return cls(**kwargs)
