"""
Tests for hanlib.database pure helpers and the in-process sqliteDatabase class.

These tests have no external database dependency (no Postgres, no Access ODBC)
and use sqlite in tmp_path so they can run in any environment.
"""
import os
import sqlite3
import sys
from datetime import date, datetime
from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest
import pandas as pd
import numpy as np

# Test against the local repo source, not the pip-installed hanetf-common.
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from hanlib import database


# =============================================================================
# castAsString
# =============================================================================

class TestCastAsString:

    def test_int(self):
        assert database.castAsString(42) == "42"

    def test_float(self):
        assert database.castAsString(3.14) == "3.14"

    def test_np_int64(self):
        assert database.castAsString(np.int64(100)) == "100"

    def test_string_passthrough(self):
        assert database.castAsString("hello") == "hello"

    def test_timestamp_uses_dmY(self):
        ts = pd.Timestamp("2025-03-09")
        assert database.castAsString(ts) == "09/03/2025"

    def test_other_falls_back_to_str(self):
        # date objects, lists, etc. all fall through to str()
        assert database.castAsString(date(2025, 3, 9)) == "2025-03-09"


# =============================================================================
# postgresDate
# =============================================================================

class TestPostgresDate:

    def test_formats_date(self):
        assert database.postgresDate(date(2025, 3, 9)) == "2025-03-09"

    def test_formats_datetime(self):
        assert database.postgresDate(datetime(2025, 3, 9, 14, 30)) == "2025-03-09"

    def test_formats_timestamp(self):
        assert database.postgresDate(pd.Timestamp("2025-12-31")) == "2025-12-31"


# =============================================================================
# booleanToString
# =============================================================================

class TestBooleanToString:

    @pytest.mark.parametrize("flavour,expected_true,expected_false", [
        ("postgres", "true", "false"),
        ("access", "-1", "0"),
        ("sqlite", 1, 0),
    ])
    def test_string_true_false_per_flavour(self, flavour, expected_true, expected_false):
        assert database.booleanToString("true", flavour) == expected_true
        assert database.booleanToString("false", flavour) == expected_false

    def test_string_case_insensitive(self):
        assert database.booleanToString("TRUE", "postgres") == "true"
        assert database.booleanToString("False", "postgres") == "false"

    def test_unknown_string_returns_false(self):
        assert database.booleanToString("maybe", "postgres") == "false"

    def test_native_true(self):
        assert database.booleanToString(True, "postgres") == "true"
        assert database.booleanToString(True, "sqlite") == 1

    def test_native_false(self):
        assert database.booleanToString(False, "postgres") == "false"
        assert database.booleanToString(False, "sqlite") == 0

    def test_truthy_value(self):
        assert database.booleanToString(1, "postgres") == "true"

    def test_falsy_value(self):
        assert database.booleanToString(0, "postgres") == "false"


# =============================================================================
# cleanHeaders
# =============================================================================

class TestCleanHeaders:

    def test_restricted_column_prefixed_sqlite(self):
        df = pd.DataFrame(columns=["date", "value", "name"])
        result = database.cleanHeaders(df, "real_navs", "sqlite")
        assert "navs_date" in result.columns
        assert "navs_value" in result.columns
        assert "name" in result.columns

    def test_postgres_strips_data_schema(self):
        df = pd.DataFrame(columns=["date", "name"])
        result = database.cleanHeaders(df, "data.navs", "postgres")
        assert "navs_date" in result.columns

    def test_postgres_strips_raw_schema(self):
        df = pd.DataFrame(columns=["date"])
        result = database.cleanHeaders(df, "raw.feed", "postgres")
        assert "feed_date" in result.columns

    def test_postgres_strips_real_schema(self):
        df = pd.DataFrame(columns=["value"])
        result = database.cleanHeaders(df, "real.index_values", "postgres")
        assert "index_values_value" in result.columns

    def test_spaces_replaced_with_underscores(self):
        df = pd.DataFrame(columns=["First Name", "Last Name"])
        result = database.cleanHeaders(df, "mytable", "sqlite")
        assert "first_name" in result.columns
        assert "last_name" in result.columns

    def test_dashes_replaced_with_underscores(self):
        df = pd.DataFrame(columns=["foo-bar", "baz-qux"])
        result = database.cleanHeaders(df, "mytable", "sqlite")
        assert "foo_bar" in result.columns
        assert "baz_qux" in result.columns

    def test_non_restricted_columns_lowercased(self):
        df = pd.DataFrame(columns=["FooBar"])
        result = database.cleanHeaders(df, "mytable", "sqlite")
        assert "foobar" in result.columns


# =============================================================================
# readData
# =============================================================================

class TestReadData:

    def test_reads_csv(self, tmp_path):
        path = tmp_path / "sample.csv"
        path.write_text("a,b\n1,2\n3,4\n")
        df = database.readData(str(path))
        assert list(df.columns) == ["a", "b"]
        assert df.shape == (2, 2)

    def test_reads_txt(self, tmp_path):
        path = tmp_path / "sample.txt"
        path.write_text("col1,col2\nfoo,bar\n")
        df = database.readData(str(path))
        assert list(df.columns) == ["col1", "col2"]
        assert df.iloc[0, 0] == "foo"

    def test_reads_excel(self, tmp_path):
        path = tmp_path / "sample.xlsx"
        pd.DataFrame({"a": [1, 2], "b": [3, 4]}).to_excel(path, index=False)
        df = database.readData(str(path))
        assert list(df.columns) == ["a", "b"]
        assert df.shape == (2, 2)

    def test_reads_excel_specific_sheet(self, tmp_path):
        path = tmp_path / "sheets.xlsx"
        with pd.ExcelWriter(path) as writer:
            pd.DataFrame({"x": [1]}).to_excel(writer, sheet_name="Sheet1", index=False)
            pd.DataFrame({"y": [9, 8]}).to_excel(writer, sheet_name="Other", index=False)
        df = database.readData(str(path), sheet_name="Other")
        assert list(df.columns) == ["y"]
        assert df.shape == (2, 1)

    def test_skiprows(self, tmp_path):
        path = tmp_path / "skip.xlsx"
        pd.DataFrame({"a": [1, 2, 3]}).to_excel(path, index=False)
        df = database.readData(str(path), skiprows=1)
        # First data row (a=1) becomes the header
        assert df.shape == (2, 1)

    def test_unsupported_extension_raises(self, tmp_path):
        path = tmp_path / "weird.parquet"
        path.write_bytes(b"not really parquet")
        with pytest.raises(Exception, match="cannot be read"):
            database.readData(str(path))


# =============================================================================
# sqliteDatabase
# =============================================================================

@pytest.fixture
def sqlite_db(tmp_path, monkeypatch):
    """Create a sqliteDatabase backed by a file in tmp_path."""
    # os.getlogin() can fail under some test runners; pin it.
    monkeypatch.setattr(os, "getlogin", lambda: "testuser")
    db = database.sqliteDatabase(db_path=str(tmp_path / "test.db"))
    yield db
    db.disconnect()


class TestSqliteDatabase:

    def test_init_sets_flavour(self, sqlite_db):
        assert sqlite_db.flavour == "sqlite"

    def test_custom_db_path_creates_file(self, tmp_path, monkeypatch):
        monkeypatch.setattr(os, "getlogin", lambda: "testuser")
        path = tmp_path / "custom.db"
        db = database.sqliteDatabase(db_path=str(path))
        try:
            assert path.exists()
        finally:
            db.disconnect()

    def test_create_table_from_dataframe(self, sqlite_db):
        df = pd.DataFrame({
            "id": [1, 2, 3],
            "name": ["alpha", "beta", "gamma"],
        })
        sqlite_db.createTable(df, "mytable")
        result = sqlite_db.query("select * from mytable order by id")
        assert result.shape == (3, 2)
        assert list(result["name"]) == ["alpha", "beta", "gamma"]

    def test_create_table_replaces_existing(self, sqlite_db):
        df1 = pd.DataFrame({"id": [1], "name": ["first"]})
        sqlite_db.createTable(df1, "mytable")
        df2 = pd.DataFrame({"id": [9], "name": ["second"]})
        sqlite_db.createTable(df2, "mytable")
        result = sqlite_db.query("select * from mytable")
        assert result.shape == (1, 2)
        assert result.iloc[0]["name"] == "second"

    def test_append_mode_adds_rows(self, sqlite_db):
        df1 = pd.DataFrame({"id": [1, 2], "name": ["a", "b"]})
        sqlite_db.createTable(df1, "mytable")
        df2 = pd.DataFrame({"id": [3, 4], "name": ["c", "d"]})
        sqlite_db.createTable(df2, "mytable", append=True)
        result = sqlite_db.query("select * from mytable order by id")
        assert result.shape == (4, 2)
        assert list(result["name"]) == ["a", "b", "c", "d"]

    def test_append_to_nonexistent_table_creates_it(self, sqlite_db):
        df = pd.DataFrame({"id": [1], "name": ["a"]})
        sqlite_db.createTable(df, "newtable", append=True)
        result = sqlite_db.query("select * from newtable")
        assert result.shape == (1, 2)

    def test_query_write_executes_dml(self, sqlite_db):
        # NOTE: sqliteDatabase doesn't register numpy adapters for sqlite3,
        # so int columns from a pandas frame land in the DB as BLOBs and
        # numeric predicates (e.g. `where id = 1`) won't match. Filter on a
        # text column instead.
        df = pd.DataFrame({"id": [1, 2], "name": ["a", "b"]})
        sqlite_db.createTable(df, "mytable")
        sqlite_db.query("delete from mytable where name = 'a'", write=True, commit=True)
        result = sqlite_db.query("select * from mytable")
        assert result.shape == (1, 2)
        assert result.iloc[0]["name"] == "b"

    def test_query_returns_dataframe(self, sqlite_db):
        df = pd.DataFrame({"id": [1, 2], "name": ["a", "b"]})
        sqlite_db.createTable(df, "mytable")
        result = sqlite_db.query("select count(*) as n from mytable")
        assert isinstance(result, pd.DataFrame)
        assert result.iloc[0]["n"] == 2

    def test_boolean_column_stored_as_integer(self, sqlite_db):
        df = pd.DataFrame({"flag": [True, False, True]})
        sqlite_db.createTable(df, "flags")
        result = sqlite_db.query("select * from flags")
        assert sorted(result["flag"].tolist()) == [0, 1, 1]

    def test_nan_handled_in_create_table(self, sqlite_db):
        df = pd.DataFrame({"id": [1, 2], "name": ["a", None]})
        sqlite_db.createTable(df, "mytable")
        result = sqlite_db.query("select * from mytable order by id")
        assert result.shape == (2, 2)

    def test_disconnect_closes_connection(self, tmp_path, monkeypatch):
        monkeypatch.setattr(os, "getlogin", lambda: "testuser")
        db = database.sqliteDatabase(db_path=str(tmp_path / "x.db"))
        db.disconnect()
        with pytest.raises(sqlite3.ProgrammingError):
            db.cursor.execute("select 1")


# =============================================================================
# getdatabaseobject (factory)
# =============================================================================

class TestGetDatabaseObject:

    def test_passes_config_to_postgresDatabase(self):
        hc = MagicMock()
        hc.cfg = {
            "db": {
                "database": "mydb",
                "host": "localhost",
                "port": "5432",
                "user": "u",
                "password": "p",
            }
        }
        with patch.object(database, "postgresDatabase") as MockPg:
            database.getdatabaseobject(hc)
            MockPg.assert_called_once_with("mydb", "localhost", "5432", "u", "p")

    def test_custom_database_key(self):
        hc = MagicMock()
        hc.cfg = {
            "warehouse": {
                "database": "wh",
                "host": "h",
                "port": "1",
                "user": "u",
                "password": "p",
            }
        }
        with patch.object(database, "postgresDatabase") as MockPg:
            database.getdatabaseobject(hc, database="warehouse")
            MockPg.assert_called_once_with("wh", "h", "1", "u", "p")
