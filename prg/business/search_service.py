"""Search and filter service for consumers and PRGs."""

from typing import List, Dict, Any, Optional


class SearchResult:
    """Result of a search operation."""

    def __init__(self):
        self.matches: List[Dict[str, Any]] = []
        self.total_count: int = 0
        self.with_expenses_count: int = 0
        self.without_expenses_count: int = 0
        self.details: List[str] = []

    def add_match(self, item: Dict[str, Any], has_expenses: bool = True):
        """Add matched item to results."""
        self.matches.append(item)
        self.total_count += 1
        if has_expenses:
            self.with_expenses_count += 1
        else:
            self.without_expenses_count += 1

    def add_detail(self, detail: str):
        """Add detail message."""
        self.details.append(detail)


class SearchService:
    """
    Handles search and filter operations.

    Provides methods for:
    - Smart search for organizations
    - Finding consumers by district/settlement
    - Finding PRG by ID
    - Filter operations
    """

    def __init__(self, validation_service=None):
        """
        Initialize search service.

        Args:
            validation_service: ValidationService instance for expense checks
        """
        self.validation_service = validation_service

    def smart_search_organizations(
        self,
        consumer_data: List[Dict[str, Any]],
        district: str,
        settlement: str,
        street_pattern: str,
        require_expenses: bool = True
    ) -> SearchResult:
        """
        Smart search for organizations by location and street name.

        Args:
            consumer_data: List of consumer dictionaries
            district: District (MO) to filter
            settlement: Settlement to filter
            street_pattern: Street name pattern to search in consumer name
            require_expenses: If True, only include consumers with expenses

        Returns:
            SearchResult with matching organizations
        """
        result = SearchResult()

        result.add_detail(f"Поиск организаций:")
        result.add_detail(f"  Район: {district}")
        result.add_detail(f"  НП: {settlement}")
        result.add_detail(f"  Улица в названии: {street_pattern}")

        for consumer in consumer_data:
            # Check if organization
            if consumer.get('type') != 'Организация':
                continue

            # Check district (case-insensitive)
            if consumer['mo'].strip().lower() != district.strip().lower():
                continue

            # Check settlement (case-insensitive)
            if consumer['settlement'].strip().lower() != settlement.strip().lower():
                continue

            # Check street in name (case-insensitive)
            if street_pattern.lower() not in consumer['name'].lower():
                continue

            # Check expenses if required
            has_expenses = self._has_expenses(consumer)
            if require_expenses and not has_expenses:
                result.add_detail(f"  Пропуск {consumer['name']} - нет расходов")
                continue

            result.add_match(consumer, has_expenses)
            result.add_detail(f"  Найдена: {consumer['name']}")

        return result

    def find_consumers_by_location(
        self,
        consumer_data: List[Dict[str, Any]],
        district: str,
        settlement: str,
        consumer_type: Optional[str] = None
    ) -> SearchResult:
        """
        Find all consumers in a specific location.

        Args:
            consumer_data: List of consumer dictionaries
            district: District (MO) to filter
            settlement: Settlement to filter
            consumer_type: Optional filter by type ('Население' or 'Организация')

        Returns:
            SearchResult with matching consumers
        """
        result = SearchResult()

        for consumer in consumer_data:
            # Check type if specified
            if consumer_type and consumer.get('type') != consumer_type:
                continue

            # Check district (case-insensitive)
            if consumer['mo'].strip().lower() != district.strip().lower():
                continue

            # Check settlement (case-insensitive)
            if consumer['settlement'].strip().lower() != settlement.strip().lower():
                continue

            has_expenses = self._has_expenses(consumer)
            result.add_match(consumer, has_expenses)

        return result

    def find_prg_by_id(
        self,
        prg_data: List[Dict[str, Any]],
        prg_id: str
    ) -> Optional[Dict[str, Any]]:
        """
        Find PRG by ID.

        Args:
            prg_data: List of PRG dictionaries
            prg_id: PRG ID to search for

        Returns:
            PRG dictionary or None if not found
        """
        for prg in prg_data:
            if prg['prg_id'] == prg_id:
                return prg
        return None

    def find_prg_by_location(
        self,
        prg_data: List[Dict[str, Any]],
        district: str,
        settlement: str
    ) -> List[Dict[str, Any]]:
        """
        Find PRGs in a specific location.

        Args:
            prg_data: List of PRG dictionaries
            district: District (MO) to filter
            settlement: Settlement to filter

        Returns:
            List of matching PRG dictionaries
        """
        matches = []

        for prg in prg_data:
            # Check district (case-insensitive)
            if prg['mo'].strip().lower() != district.strip().lower():
                continue

            # Check settlement (case-insensitive)
            if prg['settlement'].strip().lower() != settlement.strip().lower():
                continue

            matches.append(prg)

        return matches

    def get_unique_districts(self, data: List[Dict[str, Any]]) -> List[str]:
        """
        Get list of unique districts from data.

        Args:
            data: List of dictionaries with 'mo' key

        Returns:
            Sorted list of unique district names
        """
        districts = set()
        for item in data:
            mo = item.get('mo', '').strip()
            if mo:
                districts.add(mo)
        return sorted(districts)

    def get_settlements_by_district(
        self,
        data: List[Dict[str, Any]],
        district: str
    ) -> List[str]:
        """
        Get list of unique settlements for a district.

        Args:
            data: List of dictionaries with 'mo' and 'settlement' keys
            district: District to filter

        Returns:
            Sorted list of unique settlement names in the district
        """
        settlements = set()
        for item in data:
            if item.get('mo', '').strip().lower() == district.strip().lower():
                settlement = item.get('settlement', '').strip()
                if settlement:
                    settlements.add(settlement)
        return sorted(settlements)

    def get_prg_ids_by_location(
        self,
        prg_data: List[Dict[str, Any]],
        district: str,
        settlement: str
    ) -> List[str]:
        """
        Get list of PRG IDs for a specific location.

        Args:
            prg_data: List of PRG dictionaries
            district: District to filter
            settlement: Settlement to filter

        Returns:
            Sorted list of PRG IDs in the location
        """
        prg_ids = []
        for prg in prg_data:
            if (prg.get('mo', '').strip().lower() == district.strip().lower() and
                    prg.get('settlement', '').strip().lower() == settlement.strip().lower()):
                prg_id = prg.get('prg_id', '').strip()
                if prg_id:
                    prg_ids.append(prg_id)
        return sorted(prg_ids)

    def filter_consumers_by_criteria(
        self,
        consumer_data: List[Dict[str, Any]],
        has_bindings: Optional[bool] = None,
        has_expenses: Optional[bool] = None,
        consumer_type: Optional[str] = None
    ) -> SearchResult:
        """
        Filter consumers by multiple criteria.

        Args:
            consumer_data: List of consumer dictionaries
            has_bindings: If True, only with bindings; if False, only without
            has_expenses: If True, only with expenses; if False, only without
            consumer_type: Filter by type ('Население' or 'Организация')

        Returns:
            SearchResult with matching consumers
        """
        from ..data.parsers import parse_prg_bindings

        result = SearchResult()

        for consumer in consumer_data:
            # Check type
            if consumer_type and consumer.get('type') != consumer_type:
                continue

            # Check bindings
            if has_bindings is not None:
                bindings = parse_prg_bindings(consumer.get('code', ''))
                has_binding = len(bindings) > 0
                if has_binding != has_bindings:
                    continue

            # Check expenses
            consumer_has_expenses = self._has_expenses(consumer)
            if has_expenses is not None and consumer_has_expenses != has_expenses:
                continue

            result.add_match(consumer, consumer_has_expenses)

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
