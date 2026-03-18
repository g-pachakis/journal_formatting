"""
Format plugin registry.

Auto-discovers format modules in this directory.
Each module must expose FORMAT_NAME, FORMAT_SUFFIX, and build().
"""

import importlib
import pkgutil
from pathlib import Path


def get_formats():
    """Auto-discover format plugins. Returns {FORMAT_NAME: module}."""
    formats = {}
    package_dir = Path(__file__).parent

    for finder, name, ispkg in pkgutil.iter_modules([str(package_dir)]):
        module = importlib.import_module(f'.{name}', package=__package__)
        if hasattr(module, 'FORMAT_NAME') and hasattr(module, 'build'):
            formats[module.FORMAT_NAME] = module

    return formats
