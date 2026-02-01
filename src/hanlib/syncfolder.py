import os
import shutil
import logging
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
from dotenv import load_dotenv
import sys
load_dotenv()

# --- Configuration ---

# Set logging level to DEBUG to see all sync actions

def setup_logging():
    logging.basicConfig(
            # filename = os.getenv("LOG_FILE"),
            # filemode = os.getenv("LOG_MODE", "a"),
            level=os.getenv("LOG_LEVEL", "INFO"),
            format=os.getenv("LOG_FORMAT", "%(asctime)s | %(levelname)s | %(message)s"),
            datefmt=os.getenv("LOG_DATEFMT", "%Y-%m-%d %H:%M:%S")
        )


def _copy_file(source_path, destination_path):
    """
    Copy a single file from source to destination.

    Args:
        source_path (str): Full path to the source file.
        destination_path (str): Full path to the destination file.

    Returns:
        tuple: (success, filename, error_message or None)
    """
    filename = os.path.basename(source_path)
    try:
        shutil.copy2(source_path, destination_path)  # copy2 preserves metadata
        logging.debug(f"COPIED/UPDATED: {filename}")
        return (True, filename, None)
    except Exception as e:
        logging.error(f"Failed to copy {filename}: {e}")
        return (False, filename, str(e))


def synchronise_folders(source, destination, delete_missing=False, max_workers=8):
    """
    Performs a one-way synchronization from source to destination.

    Args:
        source (str): The folder to copy from.
        destination (str): The folder to copy to.
        delete_missing (bool): If True, files in dest not found in source are deleted.
        max_workers (int): Maximum number of parallel copy threads (default 8).

    Returns:
        dict: Statistics about the synchronization operation.
    """
    logging.info(f"Starting Synchronization (Delete Missing: {delete_missing}) from: {source} to: {destination}...")

    stats = {
        'files_examined': 0,
        'files_copied': 0,
        'files_skipped': 0,
        'files_deleted': 0,
        'folders_deleted': 0,
        'errors': 0
    }
    stats_lock = threading.Lock()

    source_files = set()
    dest_files = set()
    files_to_copy = []

    # --- Scan source directory using scandir for efficiency ---
    with os.scandir(source) as entries:
        for entry in entries:
            source_files.add(entry.name)
            if entry.is_file():
                stats['files_examined'] += 1
                source_path = entry.path
                destination_path = os.path.join(destination, entry.name)
                source_mtime = entry.stat().st_mtime

                # Check if file needs copying
                needs_copy = False
                if not os.path.exists(destination_path):
                    needs_copy = True
                else:
                    try:
                        dest_mtime = os.path.getmtime(destination_path)
                        if source_mtime > dest_mtime:
                            needs_copy = True
                    except OSError:
                        needs_copy = True

                if needs_copy:
                    files_to_copy.append((source_path, destination_path))
                else:
                    stats['files_skipped'] += 1

    # --- Scan destination directory using scandir ---
    try:
        with os.scandir(destination) as entries:
            for entry in entries:
                dest_files.add(entry.name)
    except FileNotFoundError:
        pass  # Destination may not have any files yet

    # --- Parallel copy phase ---
    if files_to_copy:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(_copy_file, src, dst): (src, dst)
                for src, dst in files_to_copy
            }
            for future in as_completed(futures):
                success, filename, error = future.result()
                with stats_lock:
                    if success:
                        stats['files_copied'] += 1
                    else:
                        stats['errors'] += 1

    # --- Deletion Phase (Destination Cleanup) ---
    if delete_missing:
        for filename in dest_files:
            destination_path = os.path.join(destination, filename)

            # Check if the file exists in the destination but NOT in the source
            if filename not in source_files:
                if os.path.isfile(destination_path):
                    try:
                        os.remove(destination_path)
                        logging.warning(f"DELETED: {filename} (Not found in source)")
                        stats['files_deleted'] += 1
                    except Exception as e:
                        logging.error(f"Failed to delete {filename}: {e}")
                        stats['errors'] += 1

                # Check if the folder exists in the destination but NOT in the source
                elif os.path.isdir(destination_path):
                    try:
                        shutil.rmtree(destination_path)
                        logging.warning(f"DELETED FOLDER: {filename} (Not found in source)")
                        stats['folders_deleted'] += 1
                    except Exception as e:
                        logging.error(f"Failed to delete folder {filename}: {e}")
                        stats['errors'] += 1

    logging.info(f"Synchronization complete for {source} to {destination}. "
                 f"Examined: {stats['files_examined']}, Copied: {stats['files_copied']}, "
                 f"Skipped: {stats['files_skipped']}, Deleted: {stats['files_deleted']}, "
                 f"Errors: {stats['errors']}")

    return stats


def synchronise_folders_recursive(source, destination, delete_missing=False, sync_subfolders=False, max_workers=8):
    """
    Performs synchronization from source to destination, optionally including subfolders.

    Args:
        source (str): The folder to copy from.
        destination (str): The folder to copy to.
        delete_missing (bool): If True, files in dest not found in source are deleted.
        sync_subfolders (bool): If True, recursively synchronise all subfolders.
        max_workers (int): Maximum number of parallel copy threads (default 8).

    Returns:
        dict: Aggregated statistics about the synchronization operation.
    """
    # synchronise the current folder
    total_stats = synchronise_folders(source, destination, delete_missing=delete_missing, max_workers=max_workers)

    # If sync_subfolders is enabled, process all subdirectories
    if sync_subfolders:
        try:
            with os.scandir(source) as entries:
                for entry in entries:
                    if entry.is_dir():
                        source_path = entry.path
                        destination_path = os.path.join(destination, entry.name)

                        # Create destination subfolder if it doesn't exist
                        if not os.path.exists(destination_path):
                            os.makedirs(destination_path)
                            logging.info(f"Created subdirectory: {destination_path}")

                        # Recursively synchronise the subfolder
                        sub_stats = synchronise_folders_recursive(
                            source_path, destination_path, delete_missing, sync_subfolders, max_workers
                        )

                        # Aggregate stats
                        for key in total_stats:
                            total_stats[key] += sub_stats.get(key, 0)

        except FileNotFoundError as e:
            logging.error(f"Source or destination not found: {source} {destination}: {e}")

    return total_stats


def run_task(curdate, config):
    """
    Execute folder synchronization task.

    Args:
        curdate: Current date (unused but required for geoffrey.py interface)
        config: Dictionary with source_folder, destination_folder, delete_missing,
                synchronise_folders_recursive, max_workers

    Returns:
        dict: Result with runflag, returnstatus, and report
    """
    try:
        source_folder = config.get("source_folder")
        destination_folder = config.get("destination_folder")
        delete_missing = config.get("delete_missing", False)
        sync_subfolders = config.get("synchronise_folders_recursive", False)
        max_workers = config.get("max_workers", 8)

        stats = synchronise_folders_recursive(
            source_folder, destination_folder,
            delete_missing=delete_missing,
            sync_subfolders=sync_subfolders,
            max_workers=max_workers
        )
    except FileNotFoundError as e:
        return {
            "runflag": False,
            "returnstatus": "Failure",
            "report": f"{__name__}  failed source or destination folder {source_folder} {destination_folder} do not exist: {e}",
        }

    return {
        "runflag": True,
        "returnstatus": "Success",
        "report": f"{__name__} successful - Examined: {stats['files_examined']}, "
                  f"Copied: {stats['files_copied']}, Skipped: {stats['files_skipped']}, "
                  f"Errors: {stats['errors']}",
    }


# Backward compatibility alias
runTask = run_task


def parse_args():
    """Parse command-line arguments using argparse."""
    parser = argparse.ArgumentParser(
        description="Synchronise folders from source to destination.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python syncfolder.py /source/path /dest/path
  python syncfolder.py /source/path /dest/path --delete-missing
  python syncfolder.py /source/path /dest/path --recursive --delete-missing
  python syncfolder.py /source/path /dest/path --recursive --workers 16
        """
    )
    parser.add_argument(
        "source",
        help="Source directory to copy from"
    )
    parser.add_argument(
        "destination",
        help="Destination directory to copy to"
    )
    parser.add_argument(
        "-d", "--delete-missing",
        action="store_true",
        default=False,
        help="Delete files in destination not found in source (default: false)"
    )
    parser.add_argument(
        "-r", "--recursive",
        action="store_true",
        default=False,
        help="Recursively synchronise all subfolders (default: false)"
    )
    parser.add_argument(
        "-w", "--workers",
        type=int,
        default=8,
        help="Number of parallel copy threads (default: 8)"
    )
    return parser.parse_args()


if __name__ == "__main__":
    setup_logging()
    args = parse_args()
    stats = synchronise_folders_recursive(
        args.source,
        args.destination,
        delete_missing=args.delete_missing,
        sync_subfolders=args.recursive,
        max_workers=args.workers
    )
    print(f"\nSynchronization Summary:")
    print(f"  Files examined: {stats['files_examined']}")
    print(f"  Files copied:   {stats['files_copied']}")
    print(f"  Files skipped:  {stats['files_skipped']}")
    print(f"  Files deleted:  {stats['files_deleted']}")
    print(f"  Folders deleted: {stats['folders_deleted']}")
    print(f"  Errors:         {stats['errors']}")
