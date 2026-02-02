#!/usr/bin/env python3
"""
Cleanup script for Translation Processor output folder
Deletes files older than specified retention period

Usage:
    python cleanup_output.py

Configuration:
    - RETENTION_DAYS: How many days to keep output files
    - DRY_RUN: Set to True to test without actually deleting files

Schedule this script to run daily via Windows Task Scheduler or cron.
"""
from pathlib import Path
from datetime import datetime, timedelta
import logging
import sys

# ============================================================================
# CONFIGURATION
# ============================================================================

# Folders to clean
OUTPUT_FOLDER = Path(__file__).parent / 'output'
UPLOADS_FOLDER = Path(__file__).parent / 'uploads'

# Retention periods
OUTPUT_RETENTION_DAYS = 30  # Keep output PDFs for 30 days
UPLOAD_RETENTION_DAYS = 1   # Delete orphaned uploads after 1 day

# Dry run mode (set to True to test without deleting)
DRY_RUN = False

# ============================================================================
# LOGGING SETUP
# ============================================================================

# Create logs directory if it doesn't exist
log_dir = Path(__file__).parent / 'logs'
log_dir.mkdir(exist_ok=True)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_dir / 'cleanup.log'),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)

# ============================================================================
# CLEANUP FUNCTIONS
# ============================================================================

def cleanup_folder(folder, retention_days, file_pattern='*'):
    """
    Delete files older than retention period

    Args:
        folder (Path): Folder to clean up
        retention_days (int): Number of days to keep files
        file_pattern (str): Glob pattern for files to clean (default: all files)

    Returns:
        tuple: (deleted_count, freed_space_bytes)
    """
    if not folder.exists():
        logger.warning(f"Folder does not exist: {folder}")
        return 0, 0

    cutoff_date = datetime.now() - timedelta(days=retention_days)
    deleted_count = 0
    freed_space = 0

    logger.info(f"Cleaning up {folder.name}/ - deleting files older than {cutoff_date.strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info(f"Retention period: {retention_days} days")

    # Get all files matching pattern
    files = list(folder.glob(file_pattern))

    if not files:
        logger.info(f"No files found matching pattern '{file_pattern}' in {folder.name}/")
        return 0, 0

    logger.info(f"Found {len(files)} files to check")

    # Check each file
    for file in files:
        if file.is_file():
            try:
                # Get file modification time
                file_modified = datetime.fromtimestamp(file.stat().st_mtime)
                file_size = file.stat().st_size

                # Check if file is older than cutoff
                if file_modified < cutoff_date:
                    age_days = (datetime.now() - file_modified).days

                    if DRY_RUN:
                        logger.info(f"[DRY RUN] Would delete: {file.name} (age: {age_days} days, size: {file_size / 1024 / 1024:.2f} MB)")
                        deleted_count += 1
                        freed_space += file_size
                    else:
                        try:
                            file.unlink()
                            logger.info(f"Deleted: {file.name} (age: {age_days} days, size: {file_size / 1024 / 1024:.2f} MB)")
                            deleted_count += 1
                            freed_space += file_size
                        except Exception as e:
                            logger.error(f"Failed to delete {file.name}: {e}")
                else:
                    # File is newer than cutoff, keep it
                    age_days = (datetime.now() - file_modified).days
                    logger.debug(f"Keeping: {file.name} (age: {age_days} days)")

            except Exception as e:
                logger.error(f"Error processing {file.name}: {e}")

    # Summary
    if DRY_RUN:
        logger.info(f"[DRY RUN] Would delete {deleted_count} files, freeing {freed_space / 1024 / 1024:.2f} MB")
    else:
        if deleted_count > 0:
            logger.info(f"Deleted {deleted_count} files, freed {freed_space / 1024 / 1024:.2f} MB")
        else:
            logger.info("No files needed cleanup")

    return deleted_count, freed_space


def get_folder_stats(folder, file_pattern='*'):
    """
    Get statistics about a folder

    Args:
        folder (Path): Folder to analyze
        file_pattern (str): Glob pattern for files

    Returns:
        dict: Statistics about the folder
    """
    if not folder.exists():
        return {'exists': False}

    files = list(folder.glob(file_pattern))
    total_size = sum(f.stat().st_size for f in files if f.is_file())

    return {
        'exists': True,
        'file_count': len(files),
        'total_size_mb': total_size / 1024 / 1024,
        'folder_path': str(folder)
    }


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    """Main cleanup execution"""
    logger.info("="*70)
    logger.info("Translation Processor - File Cleanup Script")
    logger.info("="*70)

    if DRY_RUN:
        logger.warning("!!! DRY RUN MODE - No files will be deleted !!!")

    # Print folder statistics before cleanup
    logger.info("\n--- Before Cleanup ---")
    output_stats = get_folder_stats(OUTPUT_FOLDER, '*.pdf')
    upload_stats = get_folder_stats(UPLOADS_FOLDER, '*.docx')

    logger.info(f"Output folder: {output_stats.get('file_count', 0)} files, "
                f"{output_stats.get('total_size_mb', 0):.2f} MB")
    logger.info(f"Upload folder: {upload_stats.get('file_count', 0)} files, "
                f"{upload_stats.get('total_size_mb', 0):.2f} MB")

    # Cleanup output folder (PDF files)
    logger.info("\n--- Cleaning Output Folder ---")
    output_deleted, output_freed = cleanup_folder(
        OUTPUT_FOLDER,
        OUTPUT_RETENTION_DAYS,
        '*.pdf'
    )

    # Cleanup uploads folder (orphaned DOCX files)
    logger.info("\n--- Cleaning Upload Folder ---")
    upload_deleted, upload_freed = cleanup_folder(
        UPLOADS_FOLDER,
        UPLOAD_RETENTION_DAYS,
        '*.docx'
    )

    # Print folder statistics after cleanup
    logger.info("\n--- After Cleanup ---")
    output_stats_after = get_folder_stats(OUTPUT_FOLDER, '*.pdf')
    upload_stats_after = get_folder_stats(UPLOADS_FOLDER, '*.docx')

    logger.info(f"Output folder: {output_stats_after.get('file_count', 0)} files, "
                f"{output_stats_after.get('total_size_mb', 0):.2f} MB")
    logger.info(f"Upload folder: {upload_stats_after.get('file_count', 0)} files, "
                f"{upload_stats_after.get('total_size_mb', 0):.2f} MB")

    # Summary
    logger.info("\n--- Cleanup Summary ---")
    total_deleted = output_deleted + upload_deleted
    total_freed = (output_freed + upload_freed) / 1024 / 1024

    if DRY_RUN:
        logger.info(f"[DRY RUN] Would delete {total_deleted} files total")
        logger.info(f"[DRY RUN] Would free {total_freed:.2f} MB disk space")
    else:
        logger.info(f"Total files deleted: {total_deleted}")
        logger.info(f"Total disk space freed: {total_freed:.2f} MB")

    logger.info("\n" + "="*70)
    logger.info("Cleanup complete")
    logger.info("="*70)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        logger.error(f"Cleanup script failed with error: {e}", exc_info=True)
        sys.exit(1)
