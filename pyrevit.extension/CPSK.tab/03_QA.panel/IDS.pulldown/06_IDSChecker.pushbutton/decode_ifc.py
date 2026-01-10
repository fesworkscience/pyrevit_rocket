# -*- coding: utf-8 -*-
"""Decode IFC PropertySet names."""
import re
import codecs
import os

def decode_ifc_hex(hex_part):
    """Decode hex-encoded Unicode."""
    result = []
    for j in range(0, len(hex_part), 4):
        if j+4 <= len(hex_part):
            code = int(hex_part[j:j+4], 16)
            result.append(chr(code))
    return ''.join(result)

# Read IFC file (pass path as argument or use default in script directory)
import sys
if len(sys.argv) > 1:
    ifc_path = sys.argv[1]
else:
    # Default: look for .ifc file in script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    ifc_path = os.path.join(script_dir, '2.ifc')
try:
    with codecs.open(ifc_path, 'r', 'utf-8') as f:
        content = f.read()
except UnicodeDecodeError:
    # Fallback: read as latin-1 (accepts any byte)
    with codecs.open(ifc_path, 'r', 'latin-1') as f:
        content = f.read()

# Find all IFCPROPERTYSET and extract names
pset_pattern = r"IFCPROPERTYSET\s*\([^,]+,[^,]+,'([^']+)'"
matches = re.findall(pset_pattern, content)

# Decode names
decoded_names = set()
for name in matches:
    # Check for hex-encoding
    if 'X2' in name:
        # Replace hex sections
        pattern = r'\\X2\\([0-9A-Fa-f]+)\\X0\\'
        decoded = re.sub(
            pattern,
            lambda m: decode_ifc_hex(m.group(1)),
            name
        )
        decoded_names.add(decoded)
    else:
        decoded_names.add(name)

# Save to file with UTF-8
script_dir = os.path.dirname(os.path.abspath(__file__))
output_path = os.path.join(script_dir, 'psets_decoded.txt')

with codecs.open(output_path, 'w', 'utf-8') as out:
    out.write('Found PropertySets (unique):\n')
    out.write('='*50 + '\n')
    for name in sorted(decoded_names):
        out.write(name + '\n')

    # Check for specific PropertySets
    target_psets = [
        u'Местоположение',
        u'Строительные параметры',
        u'Геометрические параметры',
        u'Маркировка',
        u'Идентификация'
    ]
    out.write('\n\nChecking for target PropertySets:\n')
    out.write('='*50 + '\n')
    for target in target_psets:
        found = target in decoded_names
        status = 'FOUND' if found else 'NOT FOUND'
        out.write('{}: {}\n'.format(target, status))

print('Results saved to:', output_path)
