"""Validation service for business logic."""

import pandas as pd
from typing import List, Dict, Any, Optional, Tuple
from ..data.parsers import parse_prg_bindings


class ValidationService:
    """
    Handles validation logic for PRG Pipeline Manager.

    Provides methods for:
    - Expense validation
    - Finding unbound PRGs and consumers
    - GRS validation
    """

    def has_expenses(self, consumer: Dict[str, Any]) -> bool:
        """
        Check if consumer has expense data.

        Args:
            consumer: Consumer dictionary

        Returns:
            bool: True if consumer has positive expenses
        """
        expenses = consumer.get('expenses', '')
        if not expenses or expenses == '' or expenses == 'nan' or pd.isna(expenses):
            return False

        try:
            expenses_value = float(str(expenses).replace(',', '.'))
            return expenses_value > 0
        except (ValueError, TypeError):
            return False

    def get_consumer_expenses(self, consumer: Dict[str, Any]) -> Optional[Dict[str, float]]:
        """
        Get consumer expenses (yearly and hourly).

        Args:
            consumer: Consumer dictionary

        Returns:
            Dict with keys 'yearly' and 'hourly', or None if no valid expenses
        """
        # Get yearly expenses
        yearly_raw = consumer.get('expenses', '')
        if not yearly_raw or yearly_raw == '' or yearly_raw == 'nan' or pd.isna(yearly_raw):
            return None

        try:
            yearly_str = str(yearly_raw).replace(',', '.')
            yearly_expenses = float(yearly_str)
            if yearly_expenses <= 0:
                return None
        except (ValueError, TypeError):
            return None

        # Get hourly expenses
        hourly_raw = consumer.get('hourly_expenses', '')
        hourly_expenses = None

        # Try to get hourly from Excel first
        if hourly_raw and hourly_raw != '' and hourly_raw != 'nan' and not pd.isna(hourly_raw):
            try:
                hourly_str = str(hourly_raw).replace(',', '.')
                hourly_expenses = float(hourly_str)
                if hourly_expenses <= 0:
                    hourly_expenses = None
            except (ValueError, TypeError):
                hourly_expenses = None

        # If no hourly expenses, calculate from yearly
        if hourly_expenses is None:
            hourly_expenses = yearly_expenses / 8760

        return {
            'yearly': yearly_expenses,
            'hourly': hourly_expenses
        }

    def find_unbound_prg(self, prg_data: List[Dict[str, Any]],
                        consumer_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Find PRGs without consumers in the same district and settlement.

        Args:
            prg_data: List of PRG dictionaries
            consumer_data: List of consumer dictionaries

        Returns:
            List of PRG dictionaries without matching consumers
        """
        unbound_prg = []

        for prg in prg_data:
            prg_mo = prg['mo'].strip().lower()
            prg_settlement = prg['settlement'].strip().lower()

            has_consumers = False
            for consumer in consumer_data:
                consumer_mo = consumer['mo'].strip().lower()
                consumer_settlement = consumer['settlement'].strip().lower()

                if prg_mo == consumer_mo and prg_settlement == consumer_settlement:
                    has_consumers = True
                    break

            if not has_consumers:
                unbound_prg.append(prg)

        return unbound_prg

    def find_unbound_consumers(self, consumer_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Find consumers without PRG bindings.

        Args:
            consumer_data: List of consumer dictionaries

        Returns:
            List of consumer dictionaries without bindings
        """
        unbound_consumers = []

        for consumer in consumer_data:
            bindings = parse_prg_bindings(consumer.get('code', ''))
            if not bindings:
                unbound_consumers.append(consumer)

        return unbound_consumers

    def find_consumers_without_expenses(self, consumer_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Find consumers without expense data.

        Args:
            consumer_data: List of consumer dictionaries

        Returns:
            List of consumer dictionaries without expenses
        """
        return [c for c in consumer_data if not self.has_expenses(c)]

    def validate_binding_compatibility(self, consumer: Dict[str, Any],
                                     prg: Dict[str, Any]) -> Tuple[bool, str]:
        """
        Validate if consumer can be bound to PRG (district and settlement must match).

        Args:
            consumer: Consumer dictionary
            prg: PRG dictionary

        Returns:
            Tuple of (is_valid, error_message)
        """
        consumer_mo = consumer['mo'].strip().lower()
        consumer_settlement = consumer['settlement'].strip().lower()
        prg_mo = prg['mo'].strip().lower()
        prg_settlement = prg['settlement'].strip().lower()

        if consumer_mo != prg_mo:
            return False, f"Район не совпадает: {consumer['mo']} != {prg['mo']}"

        if consumer_settlement != prg_settlement:
            return False, f"НП не совпадает: {consumer['settlement']} != {prg['settlement']}"

        return True, ""

    def check_organization_grs_mismatches(self, consumer_data: List[Dict[str, Any]],
                                         grs_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Check for GRS mismatches in organization data.

        Finds organizations where:
        - GRS ID field is empty but they have bindings
        - GRS ID doesn't match the GRS in their PRG binding

        Args:
            consumer_data: List of consumer dictionaries
            grs_data: List of GRS dictionaries

        Returns:
            List of mismatch dictionaries with keys: consumer, issue, grs_in_id, grs_in_code
        """
        mismatches = []
        grs_lookup = {grs['grs_id']: grs for grs in grs_data}

        for consumer in consumer_data:
            # Only check organizations
            if consumer.get('type') != 'Организация':
                continue

            # Skip if no bindings
            bindings = parse_prg_bindings(consumer.get('code', ''))
            if not bindings:
                continue

            grs_id = consumer.get('grs_id', '').strip()

            # Check 1: Empty GRS ID but has bindings
            if not grs_id:
                mismatches.append({
                    'consumer': consumer,
                    'issue': 'empty_grs_id',
                    'grs_in_id': '',
                    'grs_in_code': bindings[0]['grs_name'] if bindings else ''
                })
                continue

            # Check 2: GRS ID doesn't match binding
            if bindings:
                binding_grs = bindings[0]['grs_name']
                grs_record = grs_lookup.get(grs_id)
                grs_name = grs_record['grs_name'] if grs_record else grs_id

                if binding_grs != grs_name:
                    mismatches.append({
                        'consumer': consumer,
                        'issue': 'grs_mismatch',
                        'grs_in_id': grs_name,
                        'grs_in_code': binding_grs
                    })

        return mismatches

    def get_grs_name_by_id(self, grs_data: List[Dict[str, Any]], grs_id: str) -> str:
        """
        Look up GRS name by ID.

        Args:
            grs_data: List of GRS dictionaries
            grs_id: GRS ID to look up

        Returns:
            str: GRS name or fallback string
        """
        for grs in grs_data:
            if grs.get('grs_id') == grs_id:
                return grs.get('grs_name', f"ГРС {grs_id}")
        return f"ГРС {grs_id}"
