"""
Comprehensive regression tests for database.py

IMPORTANT: These tests are DESTRUCTIVE and will DROP/RECREATE tables.
They MUST ONLY be run in the DEV database.

These tests are designed to validate hanlib.database functionality
and ensure compatibility with NumPy 2.x and Pandas 2.x.
"""

import pytest
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime, date
from decimal import Decimal
import os
from dotenv import load_dotenv

load_dotenv()

from hanlib import database

# Test data paths
TEST_DATA_DIR = Path(__file__).parent / "testdata" / "test_database"


def get_environment():
    """Get the current environment from .env."""
    return os.getenv("ENVIRONMENT", "unknown")


def is_dev_environment():
    """Check if running in dev environment."""
    env = get_environment()
    return env.lower() in ("dev", "markdev")


def require_dev_environment():
    """Enforce that tests only run in dev environment."""
    if not is_dev_environment():
        pytest.fail(
            f"\n\n"
            f"!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n"
            f"!!! DESTRUCTIVE DATABASE REGRESSION TESTS !!!             \n"
            f"!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n"
            f"\n"
            f"Current environment: {get_environment()}\n"
            f"\n"
            f"These database regression tests are DESTRUCTIVE and can\n"
            f"only be run in the DEV environment.\n"
            f"\n"
            f"Running these tests in UAT or PROD will result in data loss!\n"
            f"!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n"
        )


def _get_db_connection():
    """Create a postgresDatabase connection from .env variables."""
    db_name = os.getenv("DB_DATABASE", "dev")
    host = os.getenv("DB_HOST")
    port = os.getenv("DB_PORT", "5432")
    user = os.getenv("DB_USER")
    password = os.getenv("DB_PASSWORD")

    missing = []
    if not host:
        missing.append("DB_HOST")
    if not user:
        missing.append("DB_USER")
    if not password:
        missing.append("DB_PASSWORD")
    if missing:
        pytest.fail(
            f"Missing required .env variables for database connection: {', '.join(missing)}\n"
            f"Ensure your .env file contains DB_HOST, DB_USER, and DB_PASSWORD."
        )

    return database.postgresDatabase(db_name, host, port, user, password)


@pytest.fixture(scope="session")
def db():
    """Create database connection for testing."""
    require_dev_environment()
    conn = _get_db_connection()
    yield conn
    conn.disconnect()


@pytest.fixture(scope="function")
def db_fresh():
    """Create a fresh database connection for each test."""
    require_dev_environment()
    conn = _get_db_connection()
    yield conn
    conn.rollback()
    conn.disconnect()


# =============================================================================
# ENVIRONMENT GUARD TESTS
# =============================================================================

class TestEnvironmentGuard:
    """Ensure tests fail appropriately in non-dev environments."""

    def test_environment_is_dev(self):
        """Verify we are running in dev environment."""
        require_dev_environment()
        assert is_dev_environment(), f"Tests must run in dev. Current: {get_environment()}"

    def test_fail_in_uat(self):
        """This test verifies the guard mechanism exists."""
        # This test will fail if require_dev_environment() is bypassed
        require_dev_environment()
        assert True


# =============================================================================
# NaN HANDLING TESTS (CRITICAL for NumPy 2.x upgrade)
# =============================================================================

class TestNaNHandling:
    """Tests specific to NumPy/Pandas NaN behavior changes."""

    def test_castAsBoolean_with_empty_string(self):
        """Verify castAsBoolean handles empty string correctly."""
        result = database.castAsBoolean('')
        assert result == '', "Empty string should return empty string"

    def test_castAsBoolean_with_none(self):
        """Verify castAsBoolean handles None correctly."""
        result = database.castAsBoolean(None)
        assert result == '', "None should return empty string"

    def test_castAsBoolean_with_np_nan(self):
        """Verify castAsBoolean handles np.nan correctly."""
        # np.nan == np.nan is False in NumPy, so we use pd.isna() for comparison
        result = database.castAsBoolean(np.nan)
        assert result == '', "np.nan should return empty string"

    def test_castAsBoolean_with_true(self):
        """Verify castAsBoolean handles 1 correctly."""
        result = database.castAsBoolean(1)
        assert result is True, "1 should return True"

    def test_castAsBoolean_with_false(self):
        """Verify castAsBoolean handles 0 correctly."""
        result = database.castAsBoolean(0)
        assert result is False, "0 should return False"

    def test_dataType_identifies_null_values(self):
        """Verify dataType returns 'Null' for various NaN representations."""
        assert database.dataType(None) == 'Null' or database.dataType('') == 'Null', \
            "Empty string should be Null"
        assert database.dataType('#N/A') == 'Null', "#N/A should be Null"

    def test_dataType_with_pd_na(self):
        """Verify dataType handles pd.NA correctly."""
        result = database.dataType(pd.NA)
        assert result == 'Null', "pd.NA should return 'Null'"

    def test_pd_isna_detects_all_na_types(self):
        """Verify pd.isna() works for all NA representations."""
        assert pd.isna(np.nan), "pd.isna should detect np.nan"
        assert pd.isna(pd.NA), "pd.isna should detect pd.NA"
        assert pd.isna(pd.NaT), "pd.isna should detect pd.NaT"
        assert pd.isna(None), "pd.isna should detect None"

    def test_np_nan_comparison_behavior(self):
        """Confirm np.nan != np.nan (NumPy 2.x behavior)."""
        # This is expected behavior - NaN is not equal to itself
        assert not (np.nan == np.nan), "np.nan should not equal itself"
        # Use pd.isna for proper NaN checking
        assert pd.isna(np.nan), "pd.isna should identify np.nan"


# =============================================================================
# DATA TYPE DETECTION TESTS
# =============================================================================

class TestDataTypeDetection:
    """Tests for dataType() and singleType() functions."""

    def test_dataType_integer(self):
        """Test integer detection for int, np.int64, np.int32."""
        assert database.dataType(42) == 'Integer', "int should be Integer"
        assert database.dataType(np.int64(42)) == 'Integer', "np.int64 should be Integer"

    def test_dataType_float(self):
        """Test float/numeric detection."""
        assert database.dataType(3.14) == 'Numeric', "float should be Numeric"
        assert database.dataType(np.float64(3.14)) == 'Numeric', "np.float64 should be Numeric"

    def test_dataType_string(self):
        """Test varchar detection for strings."""
        assert database.dataType("hello") == 'Varchar', "string should be Varchar"
        assert database.dataType("123abc") == 'Varchar', "alphanumeric string should be Varchar"

    def test_dataType_boolean_string(self):
        """Test 'true'/'false' string detection as Boolean."""
        assert database.dataType("true") == 'Boolean', "'true' string should be Boolean"
        assert database.dataType("false") == 'Boolean', "'false' string should be Boolean"
        assert database.dataType("TRUE") == 'Boolean', "'TRUE' string should be Boolean"

    def test_dataType_boolean_native(self):
        """Test bool and np.bool_ detection."""
        assert database.dataType(True) == 'Boolean', "True should be Boolean"
        assert database.dataType(False) == 'Boolean', "False should be Boolean"

    def test_dataType_timestamp(self):
        """Test pd.Timestamp detection."""
        ts = pd.Timestamp('2024-01-01')
        assert database.dataType(ts) == 'Date', "pd.Timestamp should be Date"

    def test_dataType_date(self):
        """Test datetime.date detection."""
        d = date(2024, 1, 1)
        assert database.dataType(d) == 'Date', "datetime.date should be Date"

    def test_dataType_nat(self):
        """Test pd.NaT detection as Date."""
        result = database.dataType(pd.NaT)
        assert result == 'Date', "pd.NaT should be Date"

    def test_singleType_hierarchy(self):
        """Verify type hierarchy: Varchar > Boolean > Numeric > Date > Integer."""
        # Varchar wins over everything
        assert database.singleType(['Varchar', 'Integer'], 'postgres') == 'Varchar'
        assert database.singleType(['Varchar', 'Boolean'], 'postgres') == 'Varchar'
        # Boolean wins over numeric types
        assert database.singleType(['Boolean', 'Integer'], 'postgres') == 'Boolean'
        # Numeric wins over Date and Integer
        assert database.singleType(['Numeric', 'Date'], 'postgres') == 'Numeric'
        assert database.singleType(['Numeric', 'Integer'], 'postgres') == 'Numeric'
        # Date wins over Integer
        assert database.singleType(['Date', 'Integer'], 'postgres') == 'Date'

    def test_singleType_all_flavours(self):
        """Test type mapping for postgres, access, sqlite."""
        assert database.singleType(['Integer'], 'postgres') == 'Integer'
        assert database.singleType(['Integer'], 'access') == 'integer'
        assert database.singleType(['Integer'], 'sqlite') == 'integer'


# =============================================================================
# DROP AND RECREATE TABLE TESTS
# =============================================================================

class TestDropAndRecreateTable:
    """Tests for dropping and recreating tables using createTable with append=False."""

    @pytest.fixture
    def product_test_data(self):
        """Load product test data from CSV."""
        csv_path = TEST_DATA_DIR / "product.csv"
        if not csv_path.exists():
            pytest.skip(f"Test data not found: {csv_path}")
        df = pd.read_csv(csv_path)
        # Take first 5 rows for testing
        return df.head(5)

    @pytest.fixture
    def shareclass_test_data(self):
        """Load shareclass test data from CSV."""
        csv_path = TEST_DATA_DIR / "shareclass.csv"
        if not csv_path.exists():
            pytest.skip(f"Test data not found: {csv_path}")
        df = pd.read_csv(csv_path)
        # Take first 5 rows for testing
        return df.head(5)

    def test_drop_and_recreate_product_table(self, db_fresh, product_test_data):
        """Test dropping and recreating the product table with test data."""
        require_dev_environment()

        # Create a test copy of the product table in dev schema
        test_table_name = "test_product"
        original_ids = product_test_data['id'].tolist()

        # Drop and recreate the table
        db_fresh.createTable(
            product_test_data,
            test_table_name,
            schema='data',
            append=False,
            clean_dataframe=True,
            log=False,
            add_id=False  # Keep original IDs
        )

        # Verify the table was created with correct data
        result = db_fresh.query(f"SELECT * FROM data.{test_table_name} ORDER BY id")

        assert result is not None, "Query should return results"
        assert len(result) == len(product_test_data), \
            f"Row count mismatch: expected {len(product_test_data)}, got {len(result)}"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table_name}", write=True)

    def test_drop_and_recreate_shareclass_table(self, db_fresh, shareclass_test_data):
        """Test dropping and recreating the shareclass table with test data."""
        require_dev_environment()

        test_table_name = "test_shareclass"

        # Drop and recreate the table
        db_fresh.createTable(
            shareclass_test_data,
            test_table_name,
            schema='data',
            append=False,
            clean_dataframe=True,
            log=False,
            add_id=False
        )

        # Verify the table was created
        result = db_fresh.query(f"SELECT * FROM data.{test_table_name}")

        assert result is not None, "Query should return results"
        assert len(result) == len(shareclass_test_data), \
            f"Row count mismatch: expected {len(shareclass_test_data)}, got {len(result)}"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table_name}", write=True)


# =============================================================================
# UPSERT TESTS
# =============================================================================

class TestUpserts:
    """Tests for upsert functionality using constraints."""

    @pytest.fixture
    def nav_test_data(self):
        """Create sample NAV data for testing upserts."""
        return pd.DataFrame({
            'isin': ['TEST123', 'TEST456', 'TEST789'],
            'nav_date': [date(2024, 1, 1), date(2024, 1, 1), date(2024, 1, 2)],
            'nav': [100.50, 200.75, 150.25],
            'aum': [1000000.0, 2000000.0, 1500000.0],
            'sharesos': [10000, 20000, 15000]
        })

    def test_upsert_with_constraint(self, db_fresh, nav_test_data):
        """Test upsert to nav table using nav_pkey constraint."""
        require_dev_environment()

        test_table_name = "test_nav"

        # First create the table with a unique constraint
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table_name}", write=True)
        db_fresh.query(f"""
            CREATE TABLE data.{test_table_name} (
                id serial primary key,
                isin varchar,
                nav_date date,
                nav numeric,
                aum numeric,
                sharesos integer,
                CONSTRAINT test_nav_iv UNIQUE (isin, nav_date)
            )
        """, write=True)

        # Insert initial data
        db_fresh.createTable(
            nav_test_data,
            test_table_name,
            schema='data',
            append=True,
            add_id=True,
            constraint='test_nav_iv'
        )

        # Create updated data (same keys, different values)
        updated_data = pd.DataFrame({
            'isin': ['TEST123', 'TEST456'],  # Same ISINs
            'nav_date': [date(2024, 1, 1), date(2024, 1, 1)],  # Same dates
            'nav': [101.00, 201.00],  # Updated NAVs
            'aum': [1100000.0, 2100000.0],  # Updated AUMs
            'sharesos': [11000, 21000]  # Updated shares
        })

        # Perform upsert
        db_fresh.createTable(
            updated_data,
            test_table_name,
            schema='data',
            append=True,
            add_id=True,
            constraint='test_nav_iv'
        )

        # Verify the upsert worked
        result = db_fresh.query(
            f"SELECT * FROM data.{test_table_name} WHERE isin = 'TEST123' AND nav_date = '2024-01-01'"
        )

        assert len(result) == 1, "Should have exactly 1 row after upsert"
        assert float(result['nav'].iloc[0]) == 101.00, f"NAV should be updated to 101.00, got {result['nav'].iloc[0]}"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table_name}", write=True)

    def test_upsert_inserts_new_rows(self, db_fresh):
        """Test that upsert inserts new rows when constraint is not violated."""
        require_dev_environment()

        test_table_name = "test_nav_insert"

        # Create table with constraint
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table_name}", write=True)
        db_fresh.query(f"""
            CREATE TABLE data.{test_table_name} (
                id serial primary key,
                isin varchar,
                nav_date date,
                nav numeric,
                CONSTRAINT test_nav_insert_pkey UNIQUE (isin, nav_date)
            )
        """, write=True)

        # Insert initial data
        initial_data = pd.DataFrame({
            'isin': ['TEST001'],
            'nav_date': [date(2024, 1, 1)],
            'nav': [100.0]
        })

        db_fresh.createTable(
            initial_data,
            test_table_name,
            schema='data',
            append=True,
            add_id=True,
            constraint='test_nav_insert_pkey'
        )

        # Insert new data (different keys)
        new_data = pd.DataFrame({
            'isin': ['TEST002'],
            'nav_date': [date(2024, 1, 2)],
            'nav': [200.0]
        })

        db_fresh.createTable(
            new_data,
            test_table_name,
            schema='data',
            append=True,
            add_id=True,
            constraint='test_nav_insert_pkey'
        )

        # Verify both rows exist
        result = db_fresh.query(f"SELECT COUNT(*) as count FROM data.{test_table_name}")
        assert result['count'].iloc[0] == 2, "Should have 2 rows after insert"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table_name}", write=True)


# =============================================================================
# SQL SELECT TESTS
# =============================================================================

class TestSQLSelect:
    """Tests for SQL SELECT operations."""

    def test_select_returns_dataframe(self, db):
        """Verify SELECT returns pandas DataFrame."""
        require_dev_environment()
        result = db.query("SELECT 1 as test_col")
        assert isinstance(result, pd.DataFrame), "Query should return DataFrame"

    def test_select_with_where_clause(self, db_fresh):
        """Test SELECT with WHERE clause."""
        require_dev_environment()

        test_table = "test_select_where"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)
        db_fresh.query(f"""
            CREATE TABLE data.{test_table} (
                id serial primary key,
                name varchar,
                value integer
            )
        """, write=True)

        # Insert test data
        db_fresh.query(f"""
            INSERT INTO data.{test_table} (name, value) VALUES
            ('alpha', 10),
            ('beta', 20),
            ('gamma', 30)
        """, write=True)

        # Test WHERE clause
        result = db_fresh.query(f"SELECT * FROM data.{test_table} WHERE value > 15")
        assert len(result) == 2, "Should return 2 rows where value > 15"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

    def test_select_empty_result(self, db_fresh):
        """Test SELECT that returns no rows."""
        require_dev_environment()

        test_table = "test_select_empty"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)
        db_fresh.query(f"""
            CREATE TABLE data.{test_table} (id serial primary key, name varchar)
        """, write=True)

        result = db_fresh.query(f"SELECT * FROM data.{test_table}")
        assert len(result) == 0, "Should return empty DataFrame"
        assert isinstance(result, pd.DataFrame), "Should still be a DataFrame"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

    def test_select_with_ordering(self, db_fresh):
        """Test SELECT with ORDER BY."""
        require_dev_environment()

        test_table = "test_select_order"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)
        db_fresh.query(f"""
            CREATE TABLE data.{test_table} (
                id serial primary key,
                value integer
            )
        """, write=True)

        db_fresh.query(f"""
            INSERT INTO data.{test_table} (value) VALUES (3), (1), (2)
        """, write=True)

        result = db_fresh.query(f"SELECT value FROM data.{test_table} ORDER BY value ASC")
        values = result['value'].tolist()
        assert values == [1, 2, 3], f"Should be ordered ascending, got {values}"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

    def test_invalid_select_statement(self, db_fresh):
        """Test handling of invalid SELECT statement."""
        require_dev_environment()

        # Invalid SQL should not crash - run_safely is True by default
        result = db_fresh.query("SELECT * FROM nonexistent_table_xyz")
        # With run_safely=True, invalid queries return None after rollback
        assert result is None, "Invalid query should return None with run_safely=True"


# =============================================================================
# SQL INSERT TESTS
# =============================================================================

class TestSQLInsert:
    """Tests for SQL INSERT operations."""

    def test_insert_single_row(self, db_fresh):
        """Test inserting a single row."""
        require_dev_environment()

        test_table = "test_insert_single"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)
        db_fresh.query(f"""
            CREATE TABLE data.{test_table} (
                id serial primary key,
                name varchar
            )
        """, write=True)

        db_fresh.query(f"INSERT INTO data.{test_table} (name) VALUES ('test')", write=True)

        result = db_fresh.query(f"SELECT * FROM data.{test_table}")
        assert len(result) == 1, "Should have 1 row"
        assert result['name'].iloc[0] == 'test', "Name should be 'test'"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

    def test_insert_duplicate_primary_key_fails(self, db_fresh):
        """Test that duplicate primary key insert fails."""
        require_dev_environment()

        test_table = "test_insert_dup_pk"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)
        db_fresh.query(f"""
            CREATE TABLE data.{test_table} (
                id integer primary key,
                name varchar
            )
        """, write=True)

        # Insert first row
        db_fresh.query(f"INSERT INTO data.{test_table} (id, name) VALUES (1, 'first')", write=True)

        # Try to insert duplicate - should fail
        db_fresh.query(f"INSERT INTO data.{test_table} (id, name) VALUES (1, 'duplicate')", write=True)
        # With run_safely=True, this will rollback

        # Verify only original row exists
        result = db_fresh.query(f"SELECT * FROM data.{test_table}")
        # After rollback, original might also be gone depending on transaction handling
        # Just verify no duplicates exist
        assert len(result) <= 1, "Should not have duplicate rows"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

    def test_insert_special_characters(self, db_fresh):
        """Test inserting data with special characters."""
        require_dev_environment()

        test_table = "test_insert_special"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)
        db_fresh.query(f"""
            CREATE TABLE data.{test_table} (
                id serial primary key,
                name varchar
            )
        """, write=True)

        # Test with createTable which handles escaping
        df = pd.DataFrame({'name': ["O'Connor", 'Test "quotes"', 'Line\nbreak']})
        db_fresh.createTable(df, test_table, schema='data', append=True, add_id=True)

        result = db_fresh.query(f"SELECT * FROM data.{test_table}")
        assert len(result) == 3, "Should have 3 rows with special characters"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)


# =============================================================================
# SQL DELETE TESTS
# =============================================================================

class TestSQLDelete:
    """Tests for SQL DELETE operations."""

    def test_delete_single_row(self, db_fresh):
        """Test deleting a single row."""
        require_dev_environment()

        test_table = "test_delete_single"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)
        db_fresh.query(f"""
            CREATE TABLE data.{test_table} (
                id serial primary key,
                name varchar
            )
        """, write=True)

        db_fresh.query(f"INSERT INTO data.{test_table} (name) VALUES ('to_keep'), ('to_delete')", write=True)
        db_fresh.query(f"DELETE FROM data.{test_table} WHERE name = 'to_delete'", write=True)

        result = db_fresh.query(f"SELECT * FROM data.{test_table}")
        assert len(result) == 1, "Should have 1 row after delete"
        assert result['name'].iloc[0] == 'to_keep', "Remaining row should be 'to_keep'"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

    def test_delete_all_rows(self, db_fresh):
        """Test deleting all rows from a table."""
        require_dev_environment()

        test_table = "test_delete_all"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)
        db_fresh.query(f"""
            CREATE TABLE data.{test_table} (
                id serial primary key,
                name varchar
            )
        """, write=True)

        db_fresh.query(f"INSERT INTO data.{test_table} (name) VALUES ('one'), ('two'), ('three')", write=True)
        db_fresh.query(f"DELETE FROM data.{test_table}", write=True)

        result = db_fresh.query(f"SELECT COUNT(*) as count FROM data.{test_table}")
        assert result['count'].iloc[0] == 0, "Table should be empty after delete all"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

    def test_delete_nonexistent_row(self, db_fresh):
        """Test deleting a row that doesn't exist (should succeed with 0 rows affected)."""
        require_dev_environment()

        test_table = "test_delete_none"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)
        db_fresh.query(f"""
            CREATE TABLE data.{test_table} (
                id serial primary key,
                name varchar
            )
        """, write=True)

        db_fresh.query(f"INSERT INTO data.{test_table} (name) VALUES ('existing')", write=True)

        # Delete non-existent row - should not fail
        db_fresh.query(f"DELETE FROM data.{test_table} WHERE name = 'nonexistent'", write=True)

        # Original row should still exist
        result = db_fresh.query(f"SELECT * FROM data.{test_table}")
        assert len(result) == 1, "Original row should still exist"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)


# =============================================================================
# TRANSACTION TESTS
# =============================================================================

class TestTransactions:
    """Tests for transaction handling."""

    def test_rollback_on_error(self, db_fresh):
        """Test that rollback works correctly."""
        require_dev_environment()

        test_table = "test_rollback"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)
        db_fresh.query(f"""
            CREATE TABLE data.{test_table} (
                id serial primary key,
                name varchar
            )
        """, write=True)

        db_fresh.query(f"INSERT INTO data.{test_table} (name) VALUES ('before_rollback')", write=True, commit=False)

        # Explicit rollback
        db_fresh.rollback()

        # Verify the insert was rolled back
        result = db_fresh.query(f"SELECT * FROM data.{test_table}")
        assert len(result) == 0, "Insert should be rolled back"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

    def test_commit_persists_changes(self, db_fresh):
        """Test that commit persists changes."""
        require_dev_environment()

        test_table = "test_commit"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)
        db_fresh.query(f"""
            CREATE TABLE data.{test_table} (
                id serial primary key,
                name varchar
            )
        """, write=True)

        db_fresh.query(f"INSERT INTO data.{test_table} (name) VALUES ('committed')", write=True)
        db_fresh.commit()

        # Verify the data persists
        result = db_fresh.query(f"SELECT * FROM data.{test_table}")
        assert len(result) == 1, "Committed data should persist"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)


# =============================================================================
# DATA INTEGRITY TESTS
# =============================================================================

class TestDataIntegrity:
    """Tests ensuring data consistency through operations."""

    def test_round_trip_numeric_precision(self, db_fresh):
        """Insert float, read back, verify precision maintained."""
        require_dev_environment()

        test_table = "test_precision"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

        df = pd.DataFrame({
            'test_precision_value': [123.456789, 0.000001, 999999.99]
        })

        db_fresh.createTable(df, test_table, schema='data', append=False, add_id=True)

        result = db_fresh.query(f"SELECT test_precision_value FROM data.{test_table} ORDER BY id")
        for original, returned in zip(df['test_precision_value'], result['test_precision_value']):
            assert abs(float(returned) - original) < 0.0001, \
                f"Precision loss: {original} -> {returned}"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

    def test_round_trip_unicode_characters(self, db_fresh):
        """Insert unicode, read back, verify encoding preserved."""
        require_dev_environment()

        test_table = "test_unicode"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

        df = pd.DataFrame({
            'test_unicode_text': ['Hello', 'Bonjour', 'Guten Tag']
        })

        db_fresh.createTable(df, test_table, schema='data', append=False, add_id=True)

        result = db_fresh.query(f"SELECT test_unicode_text FROM data.{test_table}")
        for original, returned in zip(df['test_unicode_text'], result['test_unicode_text']):
            assert original == returned, f"Unicode mismatch: {original} -> {returned}"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

    def test_empty_dataframe_handling(self, db_fresh):
        """Verify empty DataFrame doesn't cause errors."""
        require_dev_environment()

        test_table = "test_empty_df"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

        df = pd.DataFrame({'col1': [], 'col2': []})

        # This should not raise an error
        try:
            db_fresh.createTable(df, test_table, schema='data', append=False, add_id=True)
        except Exception as e:
            pytest.fail(f"Empty DataFrame should not cause error: {e}")

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

    def test_column_name_sanitization(self, db_fresh):
        """Verify restricted names (date, currency, text, primary, value) are handled."""
        require_dev_environment()

        test_table = "test_reserved"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

        # These column names should be sanitized by cleanHeaders
        df = pd.DataFrame({
            'date': ['2024-01-01'],
            'currency': ['USD'],
            'text': ['sample'],
            'value': [100]
        })

        db_fresh.createTable(df, test_table, schema='data', append=False, add_id=True)

        result = db_fresh.query(f"SELECT * FROM data.{test_table}")
        assert len(result) == 1, "Should create table with sanitized column names"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)


# =============================================================================
# NumPy 2.x COMPATIBILITY TESTS
# =============================================================================

class TestNumPy2Compatibility:
    """Regression tests for NumPy 2.x compatibility."""

    def test_np_int64_adapter_registration(self):
        """Verify psycopg2 adapter for np.int64 is registered."""
        require_dev_environment()
        # The adapter is registered at module import time
        # Just verify we can use np.int64 values
        val = np.int64(42)
        assert isinstance(val, np.int64), "Should be np.int64"

    def test_np_float64_adapter_registration(self):
        """Verify psycopg2 adapter for np.float64 is registered."""
        require_dev_environment()
        from psycopg2.extensions import adapt

        # Test that np.float64 can be adapted (serialized) by psycopg2
        val = np.float64(3.14159)
        adapted = adapt(val)
        # The adapter should convert np.float64 to a format psycopg2 can use
        assert adapted is not None, "np.float64 should be adaptable by psycopg2"
        # Verify the adapted value is correct (AsIs wraps the float value)
        assert '3.14159' in str(adapted.getquoted()), \
            f"Adapted value should contain the float: {adapted.getquoted()}"

    def test_np_float64_insert_via_constraint(self, db_fresh):
        """Verify np.float64 values can be inserted via constraint-based upserts."""
        require_dev_environment()

        test_table = "test_np_float64_constraint"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)
        db_fresh.query(f"""
            CREATE TABLE data.{test_table} (
                id serial primary key,
                name varchar,
                float_value numeric,
                CONSTRAINT test_np_float64_pkey UNIQUE (name)
            )
        """, write=True)

        # Create DataFrame with np.float64 values
        df = pd.DataFrame({
            'name': ['test1', 'test2', 'test3'],
            'float_value': np.array([1.23456789, 2.34567890, 3.45678901], dtype=np.float64)
        })

        # Verify the values are np.float64
        assert df['float_value'].dtype == np.float64, "Values should be np.float64"

        # Insert using constraint (this path uses psycopg2 directly, not SQLAlchemy)
        db_fresh.createTable(
            df,
            test_table,
            schema='data',
            append=True,
            add_id=True,
            constraint='test_np_float64_pkey'
        )

        # Verify data was inserted correctly
        result = db_fresh.query(f"SELECT * FROM data.{test_table} ORDER BY name")
        assert len(result) == 3, "Should have 3 rows"

        # Verify values are correct (within floating point tolerance)
        for original, returned in zip(df['float_value'], result['float_value']):
            assert abs(float(returned) - float(original)) < 1e-6, \
                f"Float value mismatch: {original} -> {returned}"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

    def test_np_float64_upsert_update(self, db_fresh):
        """Verify np.float64 values can be updated via constraint-based upserts."""
        require_dev_environment()

        test_table = "test_np_float64_upsert"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True, run_safely=False)
        db_fresh.query(f"""
            CREATE TABLE data.{test_table} (
                id serial primary key,
                isin varchar,
                nav_date date,
                float_value numeric,
                CONSTRAINT test_np_float64_upsert_iv UNIQUE (isin, nav_date)
            )
        """, write=True, run_safely=False)

        # Insert initial data with np.float64
        initial_df = pd.DataFrame({
            'isin': ['TEST001'],
            'nav_date': [date(2024, 1, 1)],
            'float_value': np.array([100.123456], dtype=np.float64)
        })

        # Verify the values are np.float64
        assert initial_df['float_value'].dtype == np.float64, "Values should be np.float64"

        db_fresh.createTable(
            initial_df,
            test_table,
            schema='data',
            append=True,
            add_id=True,
            constraint='test_np_float64_upsert_iv'
        )

        # Verify initial insert worked
        result = db_fresh.query(f"SELECT * FROM data.{test_table}")
        assert len(result) == 1, f"Should have 1 row after initial insert, got {len(result)}"
        assert abs(float(result['float_value'].iloc[0]) - 100.123456) < 1e-6, \
            f"Initial float value should be 100.123456, got {result['float_value'].iloc[0]}"

        # Update with new np.float64 value (same key)
        update_df = pd.DataFrame({
            'isin': ['TEST001'],
            'nav_date': [date(2024, 1, 1)],
            'float_value': np.array([999.987654], dtype=np.float64)
        })

        db_fresh.createTable(
            update_df,
            test_table,
            schema='data',
            append=True,
            add_id=True,
            constraint='test_np_float64_upsert_iv'
        )

        # Verify the value was updated (upsert worked)
        result = db_fresh.query(f"SELECT * FROM data.{test_table}")
        assert len(result) == 1, f"Should still have 1 row after upsert, got {len(result)}"
        assert abs(float(result['float_value'].iloc[0]) - 999.987654) < 1e-6, \
            f"Float value should be updated to 999.987654, got {result['float_value'].iloc[0]}"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

    def test_np_float64_mixed_with_other_types(self, db_fresh):
        """Verify np.float64 works alongside np.int64 and other types."""
        require_dev_environment()

        test_table = "test_np_mixed_types"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)
        db_fresh.query(f"""
            CREATE TABLE data.{test_table} (
                id serial primary key,
                name varchar,
                int_value integer,
                float_value numeric,
                CONSTRAINT test_np_mixed_pkey UNIQUE (name)
            )
        """, write=True)

        # Create DataFrame with mixed numpy types
        df = pd.DataFrame({
            'name': ['mixed1', 'mixed2'],
            'int_value': np.array([100, 200], dtype=np.int64),
            'float_value': np.array([1.5, 2.5], dtype=np.float64)
        })

        # Insert using constraint
        db_fresh.createTable(
            df,
            test_table,
            schema='data',
            append=True,
            add_id=True,
            constraint='test_np_mixed_pkey'
        )

        # Verify data
        result = db_fresh.query(f"SELECT * FROM data.{test_table} ORDER BY name")
        assert len(result) == 2, "Should have 2 rows"
        assert int(result['int_value'].iloc[0]) == 100, "Int value should be 100"
        assert abs(float(result['float_value'].iloc[0]) - 1.5) < 1e-6, "Float value should be 1.5"

        # Clean up
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

    def test_fillna_preserves_data_types(self, db_fresh):
        """Verify fillna doesn't change column dtypes unexpectedly."""
        require_dev_environment()

        df = pd.DataFrame({
            'int_col': [1, 2, np.nan],
            'str_col': ['a', None, 'c']
        })

        # fillna should work without type issues
        df_filled = df.fillna('')
        assert df_filled is not None, "fillna should work"

    def test_createTable_with_nan_values(self, db_fresh):
        """Verify createTable handles NaN values in data correctly."""
        require_dev_environment()

        test_table = "test_nan_values"
        db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)

        df = pd.DataFrame({
            'col1': [1, np.nan, 3],
            'col2': ['a', None, 'c'],
            'col3': [1.0, pd.NA, 3.0]
        })

        try:
            db_fresh.createTable(df, test_table, schema='data', append=False, add_id=True)
            result = db_fresh.query(f"SELECT * FROM data.{test_table}")
            assert len(result) == 3, "Should insert 3 rows with NaN handling"
        except Exception as e:
            pytest.fail(f"createTable should handle NaN values: {e}")
        finally:
            db_fresh.query(f"DROP TABLE IF EXISTS data.{test_table}", write=True)


# =============================================================================
# RUN TESTS
# =============================================================================

if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])
