"""
Version information module for OATI PDF Parser.

Provides functions to retrieve application version, git commit hash,
and build timestamp for tracking releases.
"""

import os
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Optional


APP_VERSION = "4.4"


def get_git_hash() -> str:
    """
    Get the current git commit hash (short version).
    
    Returns:
        str: Short git commit hash (7 chars) or 'unknown' if not available.
    """
    try:
        result = subprocess.run(
            ['git', 'rev-parse', '--short', 'HEAD'],
            capture_output=True,
            text=True,
            timeout=5,
            cwd=Path(__file__).parent.parent
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except (subprocess.TimeoutExpired, FileNotFoundError, Exception):
        pass
    
    return 'unknown'


def get_build_timestamp() -> str:
    """
    Get the current timestamp for build tracking.
    
    Returns:
        str: Timestamp in format 'YYYY-MM-DD HH:MM:SS'.
    """
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def load_version_from_file() -> Optional[dict]:
    """
    Load version information from version.txt file (created during build).
    
    Returns:
        dict: Version info with keys: version, git_hash, build_time, or None if file not found.
    """
    version_file = Path(__file__).parent.parent / 'version.txt'
    
    if not version_file.exists():
        return None
    
    try:
        with open(version_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            info = {}
            for line in lines:
                if '=' in line:
                    key, value = line.strip().split('=', 1)
                    info[key.lower()] = value
            return info
    except Exception:
        return None


def get_full_version_string() -> str:
    """
    Get the complete version string with git hash and build time.
    
    Priority:
    1. Use version.txt if available (from built .exe)
    2. Generate from git if in development environment
    
    Returns:
        str: Full version string, e.g., "v4.1 (abc1234) - 2025-11-18 12:00:00"
    """
    # Try to load from version.txt first (built executable)
    version_info = load_version_from_file()
    
    if version_info:
        version = version_info.get('version', APP_VERSION)
        git_hash = version_info.get('git_hash', 'unknown')
        build_time = version_info.get('build_time', 'unknown')
        return f"v{version} ({git_hash}) - {build_time}"
    
    # Development environment - generate from git
    git_hash = get_git_hash()
    return f"v{APP_VERSION} ({git_hash}) - dev"


def get_short_version_string() -> str:
    """
    Get a short version string with just version and git hash.
    
    Returns:
        str: Short version string, e.g., "v4.1 (abc1234)"
    """
    version_info = load_version_from_file()
    
    if version_info:
        version = version_info.get('version', APP_VERSION)
        git_hash = version_info.get('git_hash', 'unknown')
        return f"v{version} ({git_hash})"
    
    git_hash = get_git_hash()
    return f"v{APP_VERSION} ({git_hash})"
