# -*- coding: utf-8 -*-
"""Decode IFC PropertySet names."""
import re

def decode_ifc_hex(hex_part):
    """Decode hex-encoded Unicode."""
    result = []
    for j in range(0, len(hex_part), 4):
        if j+4 <= len(hex_part):
            code = int(hex_part[j:j+4], 16)
            result.append(chr(code))
    return ''.join(result)

# Read IFC file
ifc_path = r'C:\Users\feduloves\Documents\web\rhino_cpsk\pyrevit.extension\CPSK.tab\QA.panel\IDS.pulldown\06_IDSChecker.pushbutton\2.ifc'
with open(ifc_path, 'r', encoding='utf-8', errors='ignore') as f:
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

import codecs
import os

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
