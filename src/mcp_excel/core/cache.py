"""LRU cache for Excel files with automatic memory management."""

import os
import time
from collections import OrderedDict
from pathlib import Path
from typing import Any, Optional, Tuple

import pandas as pd
import psutil


class FileCache:
    """LRU cache for loaded Excel files with memory monitoring."""

    def __init__(
        self,
        max_size: int = 5,
        max_memory_mb: int = 1024,
        idle_timeout_seconds: int = 600,
    ) -> None:
        """Initialize file cache.

        Args:
            max_size: Maximum number of files to cache
            max_memory_mb: Maximum memory usage in MB before forced cleanup
            idle_timeout_seconds: Seconds of inactivity before cache cleanup
        """
        self._cache: OrderedDict[str, Tuple[pd.DataFrame, float]] = OrderedDict()
        self._max_size = max_size
        self._max_memory_mb = max_memory_mb
        self._idle_timeout = idle_timeout_seconds
        self._last_access = time.time()

    def _compute_cache_key(self, file_path: Path) -> str:
        """Compute cache key from file path and modification time.

        Args:
            file_path: Path to the file

        Returns:
            Cache key combining absolute path and mtime
        """
        abs_path = file_path.resolve()
        mtime = os.path.getmtime(abs_path)
        return f"{abs_path}::{mtime}"

    def _check_memory_usage(self) -> float:
        """Check current process memory usage in MB.

        Returns:
            Memory usage in megabytes
        """
        process = psutil.Process()
        return process.memory_info().rss / 1024 / 1024

    def _evict_oldest(self) -> None:
        """Remove the least recently used item from cache."""
        if self._cache:
            self._cache.popitem(last=False)

    def _cleanup_if_needed(self) -> None:
        """Perform cleanup based on memory usage and idle time."""
        current_time = time.time()

        # Check idle timeout
        if current_time - self._last_access > self._idle_timeout:
            self._cache.clear()
            return

        # Check memory usage
        memory_mb = self._check_memory_usage()
        if memory_mb > self._max_memory_mb:
            # Evict half of the cache
            items_to_remove = len(self._cache) // 2
            for _ in range(items_to_remove):
                if self._cache:
                    self._evict_oldest()

    def get(self, file_path: Path, sheet_name: Optional[str] = None) -> Optional[pd.DataFrame]:
        """Retrieve DataFrame from cache if available.

        Args:
            file_path: Absolute path to the Excel file
            sheet_name: Optional sheet name (for multi-sheet caching)

        Returns:
            Cached DataFrame or None if not in cache
        """
        cache_key = self._compute_cache_key(file_path)
        if sheet_name:
            cache_key = f"{cache_key}::{sheet_name}"

        self._cleanup_if_needed()

        if cache_key in self._cache:
            # Move to end (mark as recently used)
            df, _ = self._cache.pop(cache_key)
            self._cache[cache_key] = (df, time.time())
            self._last_access = time.time()
            return df

        return None

    def put(
        self, file_path: Path, df: pd.DataFrame, sheet_name: Optional[str] = None
    ) -> None:
        """Store DataFrame in cache.

        Args:
            file_path: Absolute path to the Excel file
            df: DataFrame to cache
            sheet_name: Optional sheet name (for multi-sheet caching)
        """
        cache_key = self._compute_cache_key(file_path)
        if sheet_name:
            cache_key = f"{cache_key}::{sheet_name}"

        # Remove if already exists (to update position)
        if cache_key in self._cache:
            del self._cache[cache_key]

        # Evict oldest if at capacity
        if len(self._cache) >= self._max_size:
            self._evict_oldest()

        self._cache[cache_key] = (df, time.time())
        self._last_access = time.time()

    def invalidate(self, file_path: Path) -> None:
        """Remove all entries for a specific file from cache.

        Args:
            file_path: Path to the file to invalidate
        """
        abs_path = str(file_path.resolve())
        keys_to_remove = [key for key in self._cache if key.startswith(abs_path)]
        for key in keys_to_remove:
            del self._cache[key]

    def clear(self) -> None:
        """Clear entire cache."""
        self._cache.clear()

    def get_stats(self) -> dict[str, Any]:
        """Get cache statistics.

        Returns:
            Dictionary with cache stats
        """
        return {
            "size": len(self._cache),
            "max_size": self._max_size,
            "memory_mb": self._check_memory_usage(),
            "max_memory_mb": self._max_memory_mb,
            "idle_seconds": time.time() - self._last_access,
        }
