"""GRS (Gas Reduction Station) data model."""

from dataclasses import dataclass


@dataclass
class GRSData:
    """
    GRS (Gas Reduction Station) reference data.

    Represents a gas reduction station with its identifier and location.
    """
    # Identity
    grs_id: str
    grs_name: str

    # Location (optional)
    mo: str = ""  # Municipal district

    # Excel metadata (for persistence)
    sheet_name: str = ""
    excel_row: int = 0
    grs_id_col: int = 0
    grs_name_col: int = 0

    def __repr__(self):
        return f"GRSData(grs_id='{self.grs_id}', name='{self.grs_name}')"

    def to_dict(self):
        """Convert to dictionary (for compatibility with existing code)."""
        return {
            'grs_id': self.grs_id,
            'grs_name': self.grs_name,
            'mo': self.mo,
            'sheet_name': self.sheet_name,
            'excel_row': self.excel_row,
            'grs_id_col': self.grs_id_col,
            'grs_name_col': self.grs_name_col,
        }

    @classmethod
    def from_dict(cls, data: dict):
        """Create from dictionary (for compatibility with existing code)."""
        return cls(**data)
