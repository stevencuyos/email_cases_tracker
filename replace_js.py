import re

with open('Index.html', 'r') as f:
    content = f.read()

# Make sure `changeThemePalette` prevents click propagation to document properly
target = r'''function togglePaletteDropdown(e) {
      if (e) e.stopPropagation();'''
replacement = r'''function togglePaletteDropdown(e) {
      if (e) e.stopPropagation();'''

if target in content:
    content = content.replace(target, replacement)
    with open('Index.html', 'w') as f:
        f.write(content)
