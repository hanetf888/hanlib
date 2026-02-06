# Changes Summary

## Modified Files

### `setup_private_package.md`
Added **Step 11: Upgrading the Package on Other Machines** — documents how to upgrade the package after pushing new commits to GitHub. Covers three scenarios: upgrading to latest on main, upgrading to a specific tagged version, and editable installs that auto-update on `git pull`.

### `tests/__init__.py`
Cleared out incorrect imports (`from . import database`, `from . import utils`) that were copied from `src/hanlib/__init__.py`. The tests folder should have an empty `__init__.py` — the old contents caused `ImportError` when running pytest.

### `.gitignore`
Added `tests/__pycache__/` to prevent compiled test bytecode from being tracked.

## New Files

### `tests/test_database.py`
Rewrote the database test suite to work with the `hanlib` package as installed:
- Changed `import database` / `import hanconfig` to `from hanlib import database`.
- Removed the `sys.path` hack that pointed at a non-existent `lib/` directory.
- Removed the dependency on `hanconfig.HanConfig()` (which does not exist in this package).
- Fixtures now connect directly using `postgresDatabase()` with env vars from `.env` (`DB_HOST`, `DB_PORT`, `DB_USER`, `DB_PASSWORD`, `DB_DATABASE`).
- All tests retain the dev-only safety guard (`require_dev_environment()`), refusing to run unless `ENVIRONMENT=dev` or `ENVIRONMENT=markdev`.

### `tests/test_utils.py`
New test suite with 65 tests covering all testable functions in `hanlib.utils`:
- `is_business_day` — all 7 days of the week
- `is_weekend` — Saturday, Sunday, weekday
- `getlastbusinessday` — offsets, weekend rollback, invalid inputs
- `get_busday_x_days_from_date` — forward/backward, weekend skipping
- `get_first_business_day_of_month` — weekday/weekend starts, all 12 months
- `getlastmonthend` — 1/2 months back, first-of-month edge case, leap year
- `getlastmonthendbd` — business day result, month ending on weekend
- `castAsString` — int, float, string, None, date
- `toDate` — string-to-date and date-to-string conversion
- `divByMillion` — positive, zero, fractional
- `percent` — basic, zero, over 100
- `getFinalTermsRegulator` — FCA, CBI, SFSA, case-insensitive, no match
- `doNothing` — returns None
- `FolderNotFound` — is an Exception subclass, can be raised
- `setborder` — borders applied to openpyxl worksheet, corner cells, invalid ranges
- `findFile` — latest/earliest file, returns date, no matching files
- `zip_and_clean_folder_with_exclusions` — creates zip, excludes files, date-based exclusion, missing folder
- `current_function_name` — returns function name string

Not tested (require Windows COM with Office installed): `mailsubjectexists`, `excel_open_force_recalc`.
