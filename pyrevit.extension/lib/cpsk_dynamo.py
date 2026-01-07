# -*- coding: utf-8 -*-
"""
Dynamo script utilities for pyRevit.

Usage:
    from cpsk_dynamo import DynamoScanner, run_dynamo_script

    # Scan scripts
    scanner = DynamoScanner(scripts_folder)
    scripts = scanner.get_all_scripts()

    # Search
    results = scanner.search_scripts("columns")

    # Run script
    success, msg = run_dynamo_script(script_path)
"""

import os
import json
import time


class DynamoScanner:
    """
    Scans and caches Dynamo scripts from folder structure.

    Optimized for 1000+ scripts with:
    - Category-based lazy loading
    - 30-second cache per category
    - Search across name and path
    """

    def __init__(self, scripts_folder):
        """
        Initialize scanner.

        Args:
            scripts_folder: Root folder containing category subfolders with .dyn files
        """
        self.scripts_folder = scripts_folder
        self._cache = {}
        self._categories = []
        self._all_scripts = []

    def scan_categories(self):
        """
        Get all category folders.

        Returns:
            List of category folder names (sorted)
        """
        if not os.path.exists(self.scripts_folder):
            return []

        categories = []
        for name in os.listdir(self.scripts_folder):
            path = os.path.join(self.scripts_folder, name)
            if os.path.isdir(path) and not name.startswith('_'):
                categories.append(name)

        self._categories = sorted(categories)
        return self._categories

    def get_scripts_in_category(self, category, include_subfolders=True):
        """
        Get all .dyn scripts in a category folder.

        Args:
            category: Category folder name
            include_subfolders: If True, scan subfolders recursively

        Returns:
            List of script dicts with keys: name, path, rel_path, category
        """
        cache_key = category

        # Check cache (valid for 30 seconds)
        if cache_key in self._cache:
            cached_time, cached_data = self._cache[cache_key]
            if time.time() - cached_time < 30:
                return cached_data

        scripts = []
        category_path = os.path.join(self.scripts_folder, category)

        if os.path.exists(category_path):
            if include_subfolders:
                for root, dirs, files in os.walk(category_path):
                    for f in files:
                        if f.endswith('.dyn'):
                            full_path = os.path.join(root, f)
                            rel_path = os.path.relpath(full_path, self.scripts_folder)
                            scripts.append({
                                'name': os.path.splitext(f)[0],
                                'path': full_path,
                                'rel_path': rel_path,
                                'category': category
                            })
            else:
                for f in os.listdir(category_path):
                    if f.endswith('.dyn'):
                        full_path = os.path.join(category_path, f)
                        rel_path = os.path.relpath(full_path, self.scripts_folder)
                        scripts.append({
                            'name': os.path.splitext(f)[0],
                            'path': full_path,
                            'rel_path': rel_path,
                            'category': category
                        })

        # Sort by name
        scripts.sort(key=lambda x: x['name'].lower())

        # Cache result
        self._cache[cache_key] = (time.time(), scripts)

        return scripts

    def get_all_scripts(self, force_rescan=False):
        """
        Get all scripts from all categories.

        Args:
            force_rescan: If True, ignore cache and rescan

        Returns:
            List of all script dicts
        """
        if self._all_scripts and not force_rescan:
            return self._all_scripts

        all_scripts = []
        for cat in self.scan_categories():
            all_scripts.extend(self.get_scripts_in_category(cat))

        self._all_scripts = all_scripts
        return all_scripts

    def search_scripts(self, query, scripts=None):
        """
        Search scripts by name/path.

        Supports multi-term search (all terms must match).

        Args:
            query: Search query (space-separated terms)
            scripts: Optional list to search in (default: all scripts)

        Returns:
            List of matching script dicts
        """
        if scripts is None:
            scripts = self.get_all_scripts()

        query = query.lower().strip()
        if not query:
            return scripts

        # Split query into terms for AND search
        terms = query.split()

        results = []
        for script in scripts:
            name_lower = script['name'].lower()
            path_lower = script['rel_path'].lower()

            # All terms must match (in name or path)
            if all(term in name_lower or term in path_lower for term in terms):
                results.append(script)

        return results

    def get_script_info(self, script_path):
        """
        Extract metadata from .dyn file.

        Args:
            script_path: Full path to .dyn file

        Returns:
            Dict with keys: description, author, name
        """
        info = {
            'description': '',
            'author': '',
            'name': ''
        }

        try:
            with open(script_path, 'r') as f:
                data = json.load(f)
                info['description'] = data.get('Description', '')
                info['author'] = data.get('Author', '')
                info['name'] = data.get('Name', '')
        except:
            pass

        return info

    def clear_cache(self):
        """Clear all cached data."""
        self._cache.clear()
        self._all_scripts = []
        self._categories = []


def run_dynamo_script(script_path, close_after=False):
    """
    Run a Dynamo script.

    Opens the script in Dynamo (uses file association).

    Args:
        script_path: Full path to .dyn file
        close_after: Not used (for future implementation)

    Returns:
        Tuple (success: bool, message: str)
    """
    try:
        import os
        os.startfile(script_path)
        return True, "Script opened in Dynamo"
    except Exception as e:
        return False, str(e)


def run_dynamo_script_headless(script_path):
    """
    Run Dynamo script in headless mode (experimental).

    Requires Dynamo Sandbox or specific Revit configuration.

    Args:
        script_path: Full path to .dyn file

    Returns:
        Tuple (success: bool, message: str)
    """
    try:
        import subprocess

        # Try to find DynamoSandbox.exe
        dynamo_paths = [
            r"C:\Program Files\Dynamo\Dynamo Core\DynamoSandbox.exe",
            r"C:\Program Files\Dynamo\Dynamo Sandbox\DynamoSandbox.exe",
        ]

        dynamo_exe = None
        for path in dynamo_paths:
            if os.path.exists(path):
                dynamo_exe = path
                break

        if dynamo_exe:
            subprocess.Popen([dynamo_exe, script_path])
            return True, "Script started in Dynamo Sandbox"
        else:
            # Fallback to file association
            os.startfile(script_path)
            return True, "Script opened in Dynamo"

    except Exception as e:
        return False, str(e)


# === YAML utilities (no external dependencies) ===

def parse_yaml_simple(filepath):
    """
    Simple YAML parser for config files.

    Supports:
    - Top-level keys with string values
    - Lists (using - prefix)
    - Simple nested dicts (one level)

    Args:
        filepath: Path to YAML file

    Returns:
        Dict with parsed data
    """
    if not os.path.exists(filepath):
        return {}

    result = {}
    current_key = None
    current_list = None
    current_dict = None

    with open(filepath, 'r') as f:
        for line in f:
            line = line.rstrip()
            if not line or line.strip().startswith('#'):
                continue

            stripped = line.lstrip()
            indent = len(line) - len(stripped)

            if indent == 0 and ':' in stripped:
                key, val = stripped.split(':', 1)
                key = key.strip()
                val = val.strip()

                if val == '' or val == '[]' or val == '{}':
                    if val == '[]':
                        result[key] = []
                        current_list = result[key]
                        current_key = key
                    elif val == '{}':
                        result[key] = {}
                        current_dict = result[key]
                        current_key = key
                    else:
                        result[key] = {}
                        current_dict = result[key]
                        current_key = key
                else:
                    result[key] = val.strip('"\'')
                    current_key = None
                    current_list = None
                    current_dict = None

            elif indent > 0 and current_key:
                if stripped.startswith('- '):
                    val = stripped[2:].strip().strip('"\'')
                    if current_list is not None:
                        current_list.append(val)
                elif ':' in stripped:
                    key, val = stripped.split(':', 1)
                    key = key.strip()
                    val = val.strip().strip('"\'')
                    if current_dict is not None:
                        current_dict[key] = val

    return result


def save_yaml_simple(filepath, data):
    """
    Save data to YAML file.

    Args:
        filepath: Path to save file
        data: Dict to save
    """
    lines = []

    for key, val in data.items():
        if isinstance(val, list):
            lines.append("{}: []".format(key) if not val else "{}:".format(key))
            for item in val:
                lines.append('  - "{}"'.format(item))
        elif isinstance(val, dict):
            lines.append("{}: {{}}".format(key) if not val else "{}:".format(key))
            for k, v in val.items():
                lines.append('  {}: "{}"'.format(k, v))
        else:
            lines.append('{}: "{}"'.format(key, val))

    with open(filepath, 'w') as f:
        f.write('\n'.join(lines))
