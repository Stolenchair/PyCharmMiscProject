"""PRG (Gas Reduction Station Pipeline) data model."""

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class PRGData:
    """
    PRG (Pipeline Gas Reduction) data structure.

    Represents a gas pipeline with its location, connections, and load data.
    """
    # Identity
    id: str
    prg_id: str
    grs_id: str

    # Location
    mo: str  # Municipal district (район)
    settlement: str  # Settlement (населенный пункт)

    # Load values (calculated from consumer bindings)
    QY_pop: float = 0.0  # Population yearly volume
    QH_pop: float = 0.0  # Population hourly rate
    QY_ind: float = 0.0  # Industry yearly volume
    QH_ind: float = 0.0  # Industry hourly rate
    Year_volume: float = 0.0  # Total yearly volume
    Max_Hour: float = 0.0  # Total maximum hourly rate

    # Excel metadata (for persistence)
    sheet_name: str = ""
    excel_row: int = 0
    qy_pop_col: int = 0
    qh_pop_col: int = 0
    qy_ind_col: int = 0
    qh_ind_col: int = 0
    year_volume_col: int = 0
    max_hour_col: int = 0

    def __repr__(self):
        return f"PRGData(prg_id='{self.prg_id}', mo='{self.mo}', settlement='{self.settlement}')"

    def to_dict(self):
        """Convert to dictionary (for compatibility with existing code)."""
        return {
            'id': self.id,
            'prg_id': self.prg_id,
            'grs_id': self.grs_id,
            'mo': self.mo,
            'settlement': self.settlement,
            'QY_pop': self.QY_pop,
            'QH_pop': self.QH_pop,
            'QY_ind': self.QY_ind,
            'QH_ind': self.QH_ind,
            'Year_volume': self.Year_volume,
            'Max_Hour': self.Max_Hour,
            'sheet_name': self.sheet_name,
            'excel_row': self.excel_row,
            'qy_pop_col': self.qy_pop_col,
            'qh_pop_col': self.qh_pop_col,
            'qy_ind_col': self.qy_ind_col,
            'qh_ind_col': self.qh_ind_col,
            'year_volume_col': self.year_volume_col,
            'max_hour_col': self.max_hour_col,
        }

    @classmethod
    def from_dict(cls, data: dict):
        """Create from dictionary (for compatibility with existing code)."""
        return cls(**data)
