"""Change tracking models for persistence."""

from dataclasses import dataclass, field
from typing import Dict, Any, Optional
from datetime import datetime


@dataclass
class Change:
    """
    Base class for tracking changes to be saved to Excel.
    """
    change_id: str
    change_type: str
    description: str
    timestamp: str = field(default_factory=lambda: str(int(datetime.now().timestamp())))

    def to_dict(self) -> Dict[str, Any]:
        """Convert change to dictionary for serialization."""
        raise NotImplementedError("Subclasses must implement to_dict")


@dataclass
class PRGLoadChange(Change):
    """
    Tracks changes to PRG load calculations.

    Used when calculate_prg_load() updates the load values for a PRG.
    """
    prg_id: str = ""
    sheet_name: str = ""
    data: Dict[str, float] = field(default_factory=dict)

    def __post_init__(self):
        if not self.change_type:
            self.change_type = 'prg_load_calculation'
        if not self.description:
            self.description = f"Подсчет нагрузки для ПРГ {self.prg_id}"

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary."""
        return {
            'type': self.change_type,
            'prg_id': self.prg_id,
            'sheet_name': self.sheet_name,
            'description': self.description,
            'data': self.data,
        }

    @classmethod
    def from_dict(cls, change_id: str, data: Dict[str, Any]) -> 'PRGLoadChange':
        """Create from dictionary."""
        return cls(
            change_id=change_id,
            change_type=data.get('type', 'prg_load_calculation'),
            description=data.get('description', ''),
            prg_id=data.get('prg_id', ''),
            sheet_name=data.get('sheet_name', ''),
            data=data.get('data', {}),
        )


@dataclass
class ConsumerBindingChange(Change):
    """
    Tracks changes to consumer PRG bindings.

    Used when binding/unbinding consumers to/from PRGs.
    """
    consumer_id: str = ""
    sheet_name: str = ""
    old_value: str = ""
    new_value: str = ""
    excel_row: int = 0
    code_col: int = 0

    def __post_init__(self):
        if not self.description:
            self.description = f"Изменение привязки потребителя {self.consumer_id}"

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary."""
        return {
            'type': self.change_type,
            'consumer_id': self.consumer_id,
            'sheet_name': self.sheet_name,
            'old_value': self.old_value,
            'new_value': self.new_value,
            'excel_row': self.excel_row,
            'code_col': self.code_col,
            'description': self.description,
        }

    @classmethod
    def from_dict(cls, change_id: str, data: Dict[str, Any]) -> 'ConsumerBindingChange':
        """Create from dictionary."""
        return cls(
            change_id=change_id,
            change_type=data.get('type', 'binding_change'),
            description=data.get('description', ''),
            consumer_id=data.get('consumer_id', ''),
            sheet_name=data.get('sheet_name', ''),
            old_value=data.get('old_value', ''),
            new_value=data.get('new_value', ''),
            excel_row=data.get('excel_row', 0),
            code_col=data.get('code_col', 0),
        )
