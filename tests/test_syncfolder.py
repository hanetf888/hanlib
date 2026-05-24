"""
Tests for syncfolder.py - robocopy-backed folder synchronisation.
Uses temporary directories to test real robocopy invocations.
"""
import pytest
import os
import time
from pathlib import Path
from unittest.mock import patch

import sys
src_path = Path(__file__).parent.parent / "src"
sys.path.insert(0, str(src_path))

from hanlib.syncfolder import (
    synchronise_folders,
    run_task,
    parse_args,
    _parse_robocopy_summary,
)


@pytest.fixture
def source_dir(tmp_path):
    """Create a temporary source directory with test files."""
    src = tmp_path / "source"
    src.mkdir()
    (src / "file1.txt").write_text("content of file 1")
    (src / "file2.txt").write_text("content of file 2")
    (src / "file3.csv").write_text("a,b,c\n1,2,3")
    return src


@pytest.fixture
def dest_dir(tmp_path):
    """Create an empty temporary destination directory."""
    dst = tmp_path / "dest"
    dst.mkdir()
    return dst


@pytest.fixture
def source_with_subdirs(tmp_path):
    """Create a source directory with subdirectories."""
    src = tmp_path / "source_recursive"
    src.mkdir()
    (src / "root_file.txt").write_text("root content")
    sub1 = src / "subdir1"
    sub1.mkdir()
    (sub1 / "sub1_file.txt").write_text("sub1 content")
    sub2 = src / "subdir2"
    sub2.mkdir()
    (sub2 / "sub2_file.txt").write_text("sub2 content")
    nested = sub1 / "nested"
    nested.mkdir()
    (nested / "nested_file.txt").write_text("nested content")
    return src


class TestParseRobocopySummary:

    def test_parses_files_and_dirs_lines(self):
        output = """
               Total    Copied   Skipped  Mismatch    FAILED    Extras
    Dirs :         3         2         1         0         0         0
   Files :        10         5         5         0         0         2
"""
        stats = _parse_robocopy_summary(output)
        assert stats['files_examined'] == 10
        assert stats['files_copied'] == 5
        assert stats['files_skipped'] == 5
        assert stats['errors'] == 0
        assert stats['files_deleted'] == 2
        assert stats['folders_deleted'] == 0

    def test_missing_summary_returns_zeros(self):
        stats = _parse_robocopy_summary("nothing useful here")
        assert all(v == 0 for v in stats.values())


class TestSynchroniseFolders:

    def test_copies_all_files(self, source_dir, dest_dir):
        stats = synchronise_folders(str(source_dir), str(dest_dir))
        assert stats['files_copied'] == 3
        assert stats['files_examined'] == 3
        assert stats['errors'] == 0
        assert (dest_dir / "file1.txt").exists()
        assert (dest_dir / "file2.txt").exists()
        assert (dest_dir / "file3.csv").exists()

    def test_skips_unchanged_files(self, source_dir, dest_dir):
        synchronise_folders(str(source_dir), str(dest_dir))
        stats = synchronise_folders(str(source_dir), str(dest_dir))
        assert stats['files_copied'] == 0
        assert stats['files_skipped'] == 3

    def test_copies_updated_files(self, source_dir, dest_dir):
        synchronise_folders(str(source_dir), str(dest_dir))
        time.sleep(1.1)  # robocopy timestamp resolution is ~2s with /FFT off
        (source_dir / "file1.txt").write_text("updated content")
        stats = synchronise_folders(str(source_dir), str(dest_dir))
        assert stats['files_copied'] >= 1
        assert (dest_dir / "file1.txt").read_text() == "updated content"

    def test_delete_missing_removes_extra_dest_files(self, source_dir, dest_dir):
        (dest_dir / "extra.txt").write_text("should be deleted")
        stats = synchronise_folders(str(source_dir), str(dest_dir), delete_missing=True)
        assert not (dest_dir / "extra.txt").exists()
        assert stats['files_deleted'] >= 1

    def test_no_delete_missing_keeps_extra_dest_files(self, source_dir, dest_dir):
        (dest_dir / "extra.txt").write_text("should remain")
        synchronise_folders(str(source_dir), str(dest_dir), delete_missing=False)
        assert (dest_dir / "extra.txt").exists()

    def test_returns_stats_dict(self, source_dir, dest_dir):
        stats = synchronise_folders(str(source_dir), str(dest_dir))
        expected_keys = ['files_examined', 'files_copied', 'files_skipped',
                         'files_deleted', 'folders_deleted', 'errors']
        for key in expected_keys:
            assert key in stats

    def test_empty_source_copies_nothing(self, tmp_path, dest_dir):
        empty_src = tmp_path / "empty_source"
        empty_src.mkdir()
        stats = synchronise_folders(str(empty_src), str(dest_dir))
        assert stats['files_copied'] == 0
        assert stats['files_examined'] == 0

    def test_missing_source_raises(self, dest_dir):
        with pytest.raises(FileNotFoundError):
            synchronise_folders("/nonexistent/source/path", str(dest_dir))

    def test_creates_missing_destination(self, source_dir, tmp_path):
        dst = tmp_path / "brand_new_dest"
        assert not dst.exists()
        stats = synchronise_folders(str(source_dir), str(dst))
        assert dst.exists()
        assert stats['files_copied'] == 3


class TestSynchroniseFoldersRecursive:

    def test_root_only_when_not_recursive(self, source_with_subdirs, dest_dir):
        synchronise_folders(str(source_with_subdirs), str(dest_dir), recursive=False)
        assert (dest_dir / "root_file.txt").exists()
        assert not (dest_dir / "subdir1").exists()

    def test_syncs_subfolders_when_recursive(self, source_with_subdirs, dest_dir):
        synchronise_folders(str(source_with_subdirs), str(dest_dir), recursive=True)
        assert (dest_dir / "root_file.txt").exists()
        assert (dest_dir / "subdir1" / "sub1_file.txt").exists()
        assert (dest_dir / "subdir2" / "sub2_file.txt").exists()

    def test_syncs_nested_subfolders(self, source_with_subdirs, dest_dir):
        synchronise_folders(str(source_with_subdirs), str(dest_dir), recursive=True)
        assert (dest_dir / "subdir1" / "nested" / "nested_file.txt").exists()

    def test_creates_missing_dest_subdirs(self, source_with_subdirs, dest_dir):
        synchronise_folders(str(source_with_subdirs), str(dest_dir), recursive=True)
        assert (dest_dir / "subdir1").is_dir()
        assert (dest_dir / "subdir2").is_dir()
        assert (dest_dir / "subdir1" / "nested").is_dir()

    def test_counts_all_files(self, source_with_subdirs, dest_dir):
        stats = synchronise_folders(str(source_with_subdirs), str(dest_dir), recursive=True)
        # root (1) + subdir1 (1) + subdir2 (1) + nested (1) = 4
        assert stats['files_copied'] == 4
        assert stats['errors'] == 0


class TestRunTask:

    def test_run_task_success(self, source_dir, dest_dir):
        config = {
            "source_folder": str(source_dir),
            "destination_folder": str(dest_dir),
            "delete_missing": False,
            "synchronise_folders_recursive": False,
            "max_workers": 4,
        }
        result = run_task("2025-01-01", config)
        assert result["runflag"] is True
        assert result["returnstatus"] == "Success"
        assert "Copied: 3" in result["report"]

    def test_run_task_missing_source_returns_failure(self, dest_dir):
        config = {
            "source_folder": "/nonexistent/source",
            "destination_folder": str(dest_dir),
            "delete_missing": False,
            "synchronise_folders_recursive": False,
        }
        result = run_task("2025-01-01", config)
        assert result["runflag"] is False
        assert result["returnstatus"] == "Failure"

    def test_run_task_with_recursive(self, source_with_subdirs, dest_dir):
        config = {
            "source_folder": str(source_with_subdirs),
            "destination_folder": str(dest_dir),
            "delete_missing": False,
            "synchronise_folders_recursive": True,
            "max_workers": 2,
        }
        result = run_task("2025-01-01", config)
        assert result["runflag"] is True
        assert (dest_dir / "subdir1" / "sub1_file.txt").exists()


class TestParseArgs:

    def test_parse_args_required(self):
        with patch('sys.argv', ['syncfolder.py', '/source', '/dest']):
            args = parse_args()
            assert args.source == '/source'
            assert args.destination == '/dest'
            assert args.delete_missing is False
            assert args.recursive is False
            assert args.workers == 8

    def test_parse_args_delete_missing(self):
        with patch('sys.argv', ['syncfolder.py', '/source', '/dest', '--delete-missing']):
            args = parse_args()
            assert args.delete_missing is True

    def test_parse_args_recursive(self):
        with patch('sys.argv', ['syncfolder.py', '/source', '/dest', '-r']):
            args = parse_args()
            assert args.recursive is True

    def test_parse_args_workers(self):
        with patch('sys.argv', ['syncfolder.py', '/source', '/dest', '-w', '16']):
            args = parse_args()
            assert args.workers == 16
