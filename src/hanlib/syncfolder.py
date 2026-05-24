import os
import logging
import argparse
import subprocess
from dotenv import load_dotenv

load_dotenv()


def setup_logging():
    logging.basicConfig(
            level=os.getenv("LOG_LEVEL", "INFO"),
            format=os.getenv("LOG_FORMAT", "%(asctime)s | %(levelname)s | %(message)s"),
            datefmt=os.getenv("LOG_DATEFMT", "%Y-%m-%d %H:%M:%S")
        )


def _parse_robocopy_summary(output):
    """
    Parse the job summary block at the end of robocopy stdout.

    The block looks like:
                   Total    Copied   Skipped  Mismatch    FAILED    Extras
        Dirs :         3         2         1         0         0         0
       Files :        10         5         5         0         0         0

    Returns a stats dict matching the existing public contract.
    """
    stats = {
        'files_examined': 0,
        'files_copied': 0,
        'files_skipped': 0,
        'files_deleted': 0,
        'folders_deleted': 0,
        'errors': 0,
    }

    for line in output.splitlines():
        stripped = line.strip()
        if stripped.startswith('Files :'):
            parts = stripped.split(':', 1)[1].split()
            if len(parts) >= 6:
                stats['files_examined'] = int(parts[0])
                stats['files_copied'] = int(parts[1])
                stats['files_skipped'] = int(parts[2])
                stats['errors'] = int(parts[4])
                stats['files_deleted'] = int(parts[5])
        elif stripped.startswith('Dirs :'):
            parts = stripped.split(':', 1)[1].split()
            if len(parts) >= 6:
                stats['folders_deleted'] = int(parts[5])
                stats['errors'] += int(parts[4])

    return stats


def synchronise_folders(source, destination, delete_missing=False, recursive=False, max_workers=8):
    """
    One-way synchronisation from source to destination using robocopy.

    Args:
        source (str): The folder to copy from.
        destination (str): The folder to copy to.
        delete_missing (bool): If True, files/folders in dest not found in source are deleted (/PURGE).
        recursive (bool): If True, include subdirectories (/E).
        max_workers (int): Number of robocopy worker threads (/MT:N).

    Returns:
        dict: {files_examined, files_copied, files_skipped, files_deleted, folders_deleted, errors}

    Raises:
        FileNotFoundError: if source does not exist.
    """
    if not os.path.isdir(source):
        raise FileNotFoundError(f"Source folder does not exist: {source}")

    os.makedirs(destination, exist_ok=True)

    logging.info(f"Starting Synchronization (Delete Missing: {delete_missing}, "
                 f"Recursive: {recursive}) from: {source} to: {destination}...")

    cmd = [
        'robocopy', source, destination,
        f'/MT:{max_workers}',
        '/R:1',
        '/W:1',
        '/NP',
        '/NDL',
        '/NFL',
    ]
    if recursive:
        cmd.append('/E')
    if delete_missing:
        cmd.append('/PURGE')

    result = subprocess.run(cmd, capture_output=True, text=True)

    # Robocopy exit codes: 0-7 = success variants, 8+ = failure
    if result.returncode >= 8:
        logging.error(f"Robocopy failed (exit {result.returncode}):\n{result.stdout}\n{result.stderr}")

    stats = _parse_robocopy_summary(result.stdout)
    if result.returncode >= 8 and stats['errors'] == 0:
        stats['errors'] = 1

    logging.info(f"Synchronization complete for {source} to {destination}. "
                 f"Examined: {stats['files_examined']}, Copied: {stats['files_copied']}, "
                 f"Skipped: {stats['files_skipped']}, Deleted: {stats['files_deleted']}, "
                 f"Errors: {stats['errors']}")

    return stats


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
    source_folder = config.get("source_folder")
    destination_folder = config.get("destination_folder")
    try:
        stats = synchronise_folders(
            source_folder,
            destination_folder,
            delete_missing=config.get("delete_missing", False),
            recursive=config.get("synchronise_folders_recursive", False),
            max_workers=config.get("max_workers", 8),
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


runTask = run_task


def parse_args():
    """Parse command-line arguments using argparse."""
    parser = argparse.ArgumentParser(
        description="Synchronise folders from source to destination (robocopy backend).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python syncfolder.py /source/path /dest/path
  python syncfolder.py /source/path /dest/path --delete-missing
  python syncfolder.py /source/path /dest/path --recursive --delete-missing
  python syncfolder.py /source/path /dest/path --recursive --workers 16
        """
    )
    parser.add_argument("source", help="Source directory to copy from")
    parser.add_argument("destination", help="Destination directory to copy to")
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
        help="Number of robocopy worker threads (default: 8)"
    )
    return parser.parse_args()


if __name__ == "__main__":
    setup_logging()
    args = parse_args()
    stats = synchronise_folders(
        args.source,
        args.destination,
        delete_missing=args.delete_missing,
        recursive=args.recursive,
        max_workers=args.workers
    )
    print(f"\nSynchronization Summary:")
    print(f"  Files examined: {stats['files_examined']}")
    print(f"  Files copied:   {stats['files_copied']}")
    print(f"  Files skipped:  {stats['files_skipped']}")
    print(f"  Files deleted:  {stats['files_deleted']}")
    print(f"  Folders deleted: {stats['folders_deleted']}")
    print(f"  Errors:         {stats['errors']}")
