"""Calculation service for PRG load computations."""

from typing import List, Dict, Any, Tuple
from ..data.parsers import parse_prg_bindings


class CalculationResult:
    """Result of PRG load calculation operation."""

    def __init__(self):
        self.prg_loads: Dict[str, Dict[str, float]] = {}
        self.processed_consumers: int = 0
        self.processed_bindings: int = 0
        self.updated_prg_count: int = 0
        self.errors: List[str] = []
        self.details: List[str] = []  # Detailed log entries

    def add_error(self, error: str):
        """Add error message to result."""
        self.errors.append(error)

    def add_detail(self, detail: str):
        """Add detail message to result."""
        self.details.append(detail)


class CalculationService:
    """
    Handles PRG load calculation logic.

    Provides methods for:
    - Calculating PRG loads from consumer bindings
    - Separating population and organization loads
    - Computing totals (Year_volume, Max_Hour)
    """

    def __init__(self, validation_service=None):
        """
        Initialize calculation service.

        Args:
            validation_service: ValidationService instance for expense retrieval
        """
        self.validation_service = validation_service

    def calculate_prg_loads(
        self,
        prg_data: List[Dict[str, Any]],
        consumer_data: List[Dict[str, Any]]
    ) -> CalculationResult:
        """
        Calculate PRG loads from consumer bindings.

        Logic:
        1. Process all consumers with expenses
        2. Extract PRG bindings from each consumer
        3. Accumulate loads by PRG ID:
           - QY_pop (population yearly volume)
           - QH_pop (population hourly rate)
           - QY_ind (organization yearly volume)
           - QH_ind (organization hourly rate)
        4. Calculate totals:
           - Year_volume = QY_pop + QY_ind
           - Max_Hour = QH_pop + QH_ind

        Args:
            prg_data: List of PRG dictionaries
            consumer_data: List of consumer dictionaries

        Returns:
            CalculationResult with prg_loads dictionary and statistics
        """
        result = CalculationResult()

        # Process each consumer
        for consumer in consumer_data:
            try:
                # Get consumer expenses
                expenses = self._get_expenses(consumer)
                if not expenses or (expenses.get('yearly', 0) == 0 and expenses.get('hourly', 0) == 0):
                    continue  # Skip consumers without expenses

                # Get consumer bindings
                bindings = parse_prg_bindings(consumer.get('code', ''))
                if not bindings:
                    continue  # Skip unbound consumers

                result.processed_consumers += 1

                # Determine consumer type
                is_population = (consumer.get('type') == 'Население')
                is_organization = (consumer.get('type') == 'Организация')

                # Process each binding
                for binding in bindings:
                    prg_id = binding['prg_id']
                    share = binding['share']

                    # Initialize PRG load if not exists
                    if prg_id not in result.prg_loads:
                        result.prg_loads[prg_id] = {
                            'QY_pop': 0.0,  # Population yearly volume
                            'QH_pop': 0.0,  # Population hourly rate
                            'QY_ind': 0.0,  # Organization yearly volume
                            'QH_ind': 0.0   # Organization hourly rate
                        }

                    # Add loads with share weighting
                    yearly_load = expenses['yearly'] * share
                    hourly_load = expenses['hourly'] * share

                    if is_population:
                        result.prg_loads[prg_id]['QY_pop'] += yearly_load
                        result.prg_loads[prg_id]['QH_pop'] += hourly_load
                    elif is_organization:
                        result.prg_loads[prg_id]['QY_ind'] += yearly_load
                        result.prg_loads[prg_id]['QH_ind'] += hourly_load

                    result.processed_bindings += 1

                    # Log detail
                    result.add_detail(
                        f"{consumer.get('name', 'Unknown')} (тип: {consumer.get('type')}) -> "
                        f"ПРГ {prg_id}: доля {share:.3f}, годовая {yearly_load:.3f}, "
                        f"часовая {hourly_load:.3f}"
                    )

            except Exception as e:
                error_msg = f"Ошибка обработки потребителя {consumer.get('name', 'Unknown')}: {str(e)}"
                result.add_error(error_msg)
                continue

        # Count updated PRGs
        result.updated_prg_count = len(result.prg_loads)

        return result

    def apply_loads_to_prg_data(
        self,
        prg_data: List[Dict[str, Any]],
        prg_loads: Dict[str, Dict[str, float]]
    ) -> int:
        """
        Apply calculated loads to PRG data structures.

        Updates PRG dictionaries with calculated load values and totals.
        PRGs without bindings are set to zero.

        Args:
            prg_data: List of PRG dictionaries to update
            prg_loads: Dictionary of calculated loads by prg_id

        Returns:
            int: Number of PRGs updated
        """
        updated_count = 0

        for prg in prg_data:
            prg_id = prg['prg_id']

            if prg_id in prg_loads:
                load = prg_loads[prg_id]

                # Update PRG load values
                prg['QY_pop'] = load['QY_pop']
                prg['QH_pop'] = load['QH_pop']
                prg['QY_ind'] = load['QY_ind']
                prg['QH_ind'] = load['QH_ind']

                # Calculate totals
                prg['Year_volume'] = load['QY_pop'] + load['QY_ind']
                prg['Max_Hour'] = load['QH_pop'] + load['QH_ind']

                updated_count += 1
            else:
                # PRG without bindings - set to zero
                prg['QY_pop'] = 0.0
                prg['QH_pop'] = 0.0
                prg['QY_ind'] = 0.0
                prg['QH_ind'] = 0.0
                prg['Year_volume'] = 0.0
                prg['Max_Hour'] = 0.0

        return updated_count

    def calculate_consumer_total_share(self, bindings: List[Dict[str, Any]]) -> float:
        """
        Calculate total share for consumer's bindings.

        Args:
            bindings: List of binding dictionaries with 'share' key

        Returns:
            float: Total share (sum of all binding shares)
        """
        return sum(binding.get('share', 0.0) for binding in bindings)

    def _get_expenses(self, consumer: Dict[str, Any]) -> Dict[str, float]:
        """
        Get consumer expenses using validation service.

        Args:
            consumer: Consumer dictionary

        Returns:
            Dict with 'yearly' and 'hourly' keys, or empty dict if no expenses
        """
        if self.validation_service:
            return self.validation_service.get_consumer_expenses(consumer) or {}

        # Fallback: basic expense extraction
        expenses = consumer.get('expenses', '')
        if not expenses or expenses == '' or expenses == 'nan':
            return {}

        try:
            yearly_str = str(expenses).replace(',', '.')
            yearly_expenses = float(yearly_str)
            if yearly_expenses <= 0:
                return {}

            # Get hourly expenses or calculate from yearly
            hourly_raw = consumer.get('hourly_expenses', '')
            if hourly_raw and hourly_raw != '' and hourly_raw != 'nan':
                try:
                    hourly_str = str(hourly_raw).replace(',', '.')
                    hourly_expenses = float(hourly_str)
                    if hourly_expenses <= 0:
                        hourly_expenses = yearly_expenses / 8760
                except (ValueError, TypeError):
                    hourly_expenses = yearly_expenses / 8760
            else:
                hourly_expenses = yearly_expenses / 8760

            return {
                'yearly': yearly_expenses,
                'hourly': hourly_expenses
            }
        except (ValueError, TypeError):
            return {}
