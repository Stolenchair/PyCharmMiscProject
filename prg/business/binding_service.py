"""Binding service for PRG-consumer binding operations."""

from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime
from ..data.parsers import parse_prg_bindings, format_prg_bindings, calculate_total_share


class BindingResult:
    """Result of a binding operation."""

    def __init__(self, operation_type: str = "binding"):
        self.operation_type = operation_type
        self.success_count: int = 0
        self.skipped_count: int = 0
        self.already_bound_count: int = 0
        self.failed_consumers: List[Dict[str, Any]] = []
        self.changes: List[Dict[str, Any]] = []
        self.errors: List[str] = []
        self.details: List[str] = []

    def add_success(self, consumer: Dict[str, Any], change: Dict[str, Any]):
        """Record successful binding."""
        self.success_count += 1
        self.changes.append(change)

    def add_skip(self, consumer: Dict[str, Any], reason: str):
        """Record skipped consumer."""
        self.skipped_count += 1
        self.details.append(f"Пропущен {consumer.get('name', 'Unknown')}: {reason}")

    def add_already_bound(self, consumer: Dict[str, Any]):
        """Record already bound consumer."""
        self.already_bound_count += 1

    def add_error(self, consumer: Dict[str, Any], error: str):
        """Record error."""
        self.failed_consumers.append(consumer)
        self.errors.append(f"{consumer.get('name', 'Unknown')}: {error}")


class BindingService:
    """
    Handles PRG-consumer binding operations.

    Provides methods for:
    - Binding PRG to settlement (all consumers)
    - Binding single consumer
    - Unbinding operations
    - Binding validation
    """

    def __init__(self, validation_service=None):
        """
        Initialize binding service.

        Args:
            validation_service: ValidationService instance for validation
        """
        self.validation_service = validation_service

    def bind_prg_to_settlement(
        self,
        prg: Dict[str, Any],
        target_consumer: Dict[str, Any],
        all_consumers: List[Dict[str, Any]],
        grs_name: str,
        share: float
    ) -> BindingResult:
        """
        Bind PRG to all consumers in the same settlement with expense validation.

        Args:
            prg: PRG dictionary
            target_consumer: Selected consumer (defines target settlement)
            all_consumers: List of all consumers
            grs_name: GRS name for the PRG
            share: Share to assign to each consumer

        Returns:
            BindingResult with success/skip/error counts and changes
        """
        result = BindingResult(operation_type="settlement_bind")

        prg_id = prg['prg_id']
        target_mo = target_consumer['mo'].strip()
        target_settlement = target_consumer['settlement'].strip()

        # Find all consumers in the same settlement
        consumers_in_settlement = []
        for consumer in all_consumers:
            if (consumer['mo'].strip().lower() == target_mo.lower() and
                    consumer['settlement'].strip().lower() == target_settlement.lower()):
                consumers_in_settlement.append(consumer)

        # Categorize consumers
        for consumer in consumers_in_settlement:
            current_bindings = parse_prg_bindings(consumer.get('code', ''))

            # Check if already bound to this PRG
            already_bound = any(b['prg_id'] == prg_id for b in current_bindings)
            if already_bound:
                result.add_already_bound(consumer)
                continue

            # Check if has expenses
            if not self._has_expenses(consumer):
                result.add_skip(consumer, "нет расходов")
                continue

            # Check available share
            current_total = calculate_total_share(current_bindings)
            available_share = min(share, 1.0 - current_total)

            if available_share <= 0.001:
                result.add_skip(consumer, "недостаточно свободной доли")
                continue

            # Create binding
            try:
                new_binding = {
                    'prg_id': prg_id,
                    'share': available_share,
                    'grs_name': grs_name
                }

                current_bindings.append(new_binding)
                new_binding_string = format_prg_bindings(current_bindings)

                # Create change record
                old_code = consumer.get('code', '')
                consumer['code'] = new_binding_string

                change_id = f"settlement_bind_{consumer['id']}_{datetime.now().timestamp()}"
                change = {
                    'change_id': change_id,
                    'type': 'settlement_bind',
                    'consumer_id': consumer['id'],
                    'sheet_name': consumer['sheet_name'],
                    'row': consumer['excel_row'],
                    'col': consumer['code_col'],
                    'new_value': new_binding_string,
                    'old_value': old_code,
                    'description': f"Привязка НП: {consumer['name']} → ПРГ {prg_id}"
                }

                result.add_success(consumer, change)

            except Exception as e:
                result.add_error(consumer, str(e))

        return result

    def bind_single_consumer(
        self,
        consumer: Dict[str, Any],
        prg: Dict[str, Any],
        grs_name: str,
        share: float,
        force: bool = False
    ) -> BindingResult:
        """
        Bind single consumer to PRG.

        Args:
            consumer: Consumer dictionary
            prg: PRG dictionary
            grs_name: GRS name for the PRG
            share: Share to assign
            force: If True, skip expense validation (manual binding)

        Returns:
            BindingResult with success/error and changes
        """
        result = BindingResult(operation_type="single_bind")

        prg_id = prg['prg_id']

        try:
            # Validate expenses unless forced
            if not force and not self._has_expenses(consumer):
                result.add_skip(consumer, "нет расходов")
                return result

            # Get current bindings
            current_bindings = parse_prg_bindings(consumer.get('code', ''))

            # Check if already bound to this PRG
            existing_binding_idx = None
            for idx, binding in enumerate(current_bindings):
                if binding['prg_id'] == prg_id:
                    existing_binding_idx = idx
                    break

            # Update or add binding
            if existing_binding_idx is not None:
                # Update existing binding
                old_share = current_bindings[existing_binding_idx]['share']
                current_bindings[existing_binding_idx]['share'] = share
                current_bindings[existing_binding_idx]['grs_name'] = grs_name
                operation = "update"
            else:
                # Add new binding
                new_binding = {
                    'prg_id': prg_id,
                    'share': share,
                    'grs_name': grs_name
                }
                current_bindings.append(new_binding)
                operation = "add"

            # Format and save
            new_binding_string = format_prg_bindings(current_bindings)
            old_code = consumer.get('code', '')
            consumer['code'] = new_binding_string

            # Create change record
            change_id = f"manual_bind_{consumer['id']}_{datetime.now().timestamp()}"
            change = {
                'change_id': change_id,
                'type': 'manual_bind' if force else 'single_bind',
                'consumer_id': consumer['id'],
                'sheet_name': consumer['sheet_name'],
                'row': consumer['excel_row'],
                'col': consumer['code_col'],
                'new_value': new_binding_string,
                'old_value': old_code,
                'description': f"Привязка: {consumer['name']} → ПРГ {prg_id} (доля: {share:.3f})"
            }

            result.add_success(consumer, change)

        except Exception as e:
            result.add_error(consumer, str(e))

        return result

    def unbind_single_consumer(
        self,
        consumer: Dict[str, Any]
    ) -> BindingResult:
        """
        Remove all PRG bindings from a single consumer.

        Args:
            consumer: Consumer dictionary

        Returns:
            BindingResult with changes
        """
        result = BindingResult(operation_type="unbind")

        try:
            bindings = parse_prg_bindings(consumer.get('code', ''))
            if not bindings:
                result.add_skip(consumer, "нет привязок")
                return result

            # Clear bindings
            old_code = consumer['code']
            consumer['code'] = ''

            # Create change record
            change_id = f"unbind_{consumer['id']}_{datetime.now().timestamp()}"
            change = {
                'change_id': change_id,
                'type': 'unbind',
                'consumer_id': consumer['id'],
                'sheet_name': consumer['sheet_name'],
                'row': consumer['excel_row'],
                'col': consumer['code_col'],
                'new_value': '',
                'old_value': old_code,
                'description': f"Отвязка всех ПРГ от {consumer['name']}"
            }

            result.add_success(consumer, change)

        except Exception as e:
            result.add_error(consumer, str(e))

        return result

    def unbind_entire_settlement(
        self,
        target_consumer: Dict[str, Any],
        all_consumers: List[Dict[str, Any]]
    ) -> BindingResult:
        """
        Remove all PRG bindings from all consumers in a settlement.

        Args:
            target_consumer: Selected consumer (defines target settlement)
            all_consumers: List of all consumers

        Returns:
            BindingResult with changes
        """
        result = BindingResult(operation_type="unbind_settlement")

        target_mo = target_consumer['mo'].strip()
        target_settlement = target_consumer['settlement'].strip()

        # Find all consumers in the same settlement
        for consumer in all_consumers:
            if (consumer['mo'].strip().lower() == target_mo.lower() and
                    consumer['settlement'].strip().lower() == target_settlement.lower()):

                # Unbind this consumer
                unbind_result = self.unbind_single_consumer(consumer)

                # Merge results
                result.success_count += unbind_result.success_count
                result.skipped_count += unbind_result.skipped_count
                result.changes.extend(unbind_result.changes)
                result.errors.extend(unbind_result.errors)
                result.failed_consumers.extend(unbind_result.failed_consumers)

        return result

    def remove_prg_from_consumer(
        self,
        consumer: Dict[str, Any],
        prg_id: str
    ) -> BindingResult:
        """
        Remove specific PRG binding from consumer (keep other bindings).

        Args:
            consumer: Consumer dictionary
            prg_id: PRG ID to remove

        Returns:
            BindingResult with changes
        """
        result = BindingResult(operation_type="remove_prg")

        try:
            current_bindings = parse_prg_bindings(consumer.get('code', ''))

            # Find and remove the binding
            new_bindings = [b for b in current_bindings if b['prg_id'] != prg_id]

            if len(new_bindings) == len(current_bindings):
                result.add_skip(consumer, f"не привязан к ПРГ {prg_id}")
                return result

            # Update bindings
            new_binding_string = format_prg_bindings(new_bindings)
            old_code = consumer.get('code', '')
            consumer['code'] = new_binding_string

            # Create change record
            change_id = f"remove_prg_{consumer['id']}_{datetime.now().timestamp()}"
            change = {
                'change_id': change_id,
                'type': 'remove_prg',
                'consumer_id': consumer['id'],
                'sheet_name': consumer['sheet_name'],
                'row': consumer['excel_row'],
                'col': consumer['code_col'],
                'new_value': new_binding_string,
                'old_value': old_code,
                'description': f"Удаление привязки ПРГ {prg_id} от {consumer['name']}"
            }

            result.add_success(consumer, change)

        except Exception as e:
            result.add_error(consumer, str(e))

        return result

    def _has_expenses(self, consumer: Dict[str, Any]) -> bool:
        """
        Check if consumer has expenses (using validation service if available).

        Args:
            consumer: Consumer dictionary

        Returns:
            bool: True if consumer has expenses
        """
        if self.validation_service:
            return self.validation_service.has_expenses(consumer)

        # Fallback: basic check
        expenses = consumer.get('expenses', '')
        if not expenses or expenses == '' or expenses == 'nan':
            return False

        try:
            expenses_value = float(str(expenses).replace(',', '.'))
            return expenses_value > 0
        except (ValueError, TypeError):
            return False
