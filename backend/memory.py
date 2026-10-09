"""
memory.py — give freed memory back to the operating system.

Why this exists: building a PDF from a photo-heavy workbook briefly needs a lot
of RAM (every photo is decoded). Python frees it afterwards, but glibc's
allocator keeps the freed blocks in the process instead of returning them to
the OS. The process's memory therefore only ratchets up, request after request,
until the app restarts — and a hosting platform bills for that held memory.

release_memory() collects garbage and asks glibc to hand the free blocks back.
It is cheap (milliseconds) and harmless: on a platform without glibc's
malloc_trim (macOS, Windows, Alpine/musl) it quietly does nothing.

Pairs with two environment variables that fix the same thing at the allocator
level — see the "Memory" section of ARCHITECTURE.md / README:
    MALLOC_MMAP_THRESHOLD_=131072   MALLOC_ARENA_MAX=2
"""

import ctypes
import functools
import gc

try:
    _malloc_trim = ctypes.CDLL("libc.so.6").malloc_trim
except (OSError, AttributeError):
    _malloc_trim = None


def release_memory() -> None:
    gc.collect()
    if _malloc_trim is not None:
        _malloc_trim(0)


def releases_memory(func):
    """Decorator for the heavy (workbook/photo/PDF) endpoints: trim memory when
    the request finishes, whether it succeeded or raised. Preserves the
    function's signature so FastAPI still sees its parameters."""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        finally:
            release_memory()
    return wrapper
