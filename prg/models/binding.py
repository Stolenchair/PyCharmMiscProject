"""PRG Binding data model."""

from dataclasses import dataclass
from typing import List


@dataclass
class PRGBinding:
    """
    Represents a consumer's binding to a PRG pipeline.

    Format in Excel: "PRG_ID:share:GRS_name" (semicolon-separated for multiple)
    Example: "PRG-001:0.5:GRS_South;PRG-002:0.3:GRS_North"
    """
    prg_id: str
    share: float  # 0.0 to 1.0 (percentage of consumer expenses allocated to this PRG)
    grs_name: str = ""

    def __repr__(self):
        return f"PRGBinding(prg_id='{self.prg_id}', share={self.share:.3f})"

    def to_string(self) -> str:
        """
        Convert binding to Excel string format.

        Returns:
            str: "PRG_ID:share:GRS_name" or "PRG_ID:share" if no GRS
        """
        if self.grs_name:
            return f"{self.prg_id}:{self.share}:{self.grs_name}"
        return f"{self.prg_id}:{self.share}"

    def to_dict(self):
        """Convert to dictionary."""
        return {
            'prg_id': self.prg_id,
            'share': self.share,
            'grs_name': self.grs_name,
        }

    @classmethod
    def from_string(cls, binding_str: str) -> 'PRGBinding':
        """
        Parse binding from Excel string format.

        Args:
            binding_str: "PRG_ID:share" or "PRG_ID:share:GRS_name"

        Returns:
            PRGBinding instance

        Raises:
            ValueError: If format is invalid
        """
        parts = binding_str.split(':')

        if len(parts) < 2:
            raise ValueError(f"Invalid binding format: {binding_str}")

        prg_id = parts[0].strip()

        try:
            share = float(parts[1].strip())
        except ValueError:
            raise ValueError(f"Invalid share value in binding: {binding_str}")

        grs_name = parts[2].strip() if len(parts) > 2 else ""

        return cls(prg_id=prg_id, share=share, grs_name=grs_name)

    @classmethod
    def parse_bindings(cls, code: str) -> List['PRGBinding']:
        """
        Parse multiple bindings from consumer code.

        Args:
            code: Semicolon-separated binding string from Excel

        Returns:
            List of PRGBinding instances
        """
        if not code or not isinstance(code, str) or code.strip() == '':
            return []

        bindings = []
        for binding_str in code.split(';'):
            binding_str = binding_str.strip()
            if binding_str:
                try:
                    binding = cls.from_string(binding_str)
                    bindings.append(binding)
                except ValueError:
                    # Skip invalid bindings
                    continue

        return bindings

    @classmethod
    def format_bindings(cls, bindings: List['PRGBinding']) -> str:
        """
        Format multiple bindings to Excel string.

        Args:
            bindings: List of PRGBinding instances

        Returns:
            str: Semicolon-separated binding string
        """
        return ';'.join(b.to_string() for b in bindings)


def calculate_total_share(bindings: List[PRGBinding]) -> float:
    """
    Calculate total share from list of bindings.

    Args:
        bindings: List of PRGBinding instances

    Returns:
        float: Sum of all shares
    """
    return sum(b.share for b in bindings)
