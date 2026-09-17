import re

with open('Index.html', 'r') as f:
    content = f.read()

# Make the palette dropdown match the requested custom Tailwind styling better, replacing bg-white with bg-theme-sidebar
# and hover:bg-gray-100 with hover:bg-theme-canvas so it inherits the correct theme variables.

target = r'''<div id="paletteDropdown" class="hidden absolute bottom-8 left-0 mt-2 w-32 bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 rounded-theme shadow-lg overflow-hidden z-50 transition-opacity">
            <ul class="py-1 text-sm text-gray-700 dark:text-gray-200">
              <li>
                <button onclick="changeThemePalette('default')" class="w-full text-left px-4 py-2 hover:bg-gray-100 dark:hover:bg-gray-700">Default</button>
              </li>
              <li>
                <button onclick="changeThemePalette('fintech')" class="w-full text-left px-4 py-2 hover:bg-gray-100 dark:hover:bg-gray-700">Fintech</button>
              </li>
              <li>
                <button onclick="changeThemePalette('hig')" class="w-full text-left px-4 py-2 hover:bg-gray-100 dark:hover:bg-gray-700">macOS</button>
              </li>
            </ul>
          </div>'''

replacement = r'''<div id="paletteDropdown" class="hidden absolute bottom-8 left-0 mt-2 w-36 bg-theme-sidebar border border-theme-card-border rounded-theme shadow-theme overflow-hidden z-50 transition-opacity">
            <ul class="py-1 text-sm text-gray-700 dark:text-gray-200">
              <li>
                <button onclick="changeThemePalette('default')" class="flex items-center gap-2 w-full text-left px-4 py-2 hover:bg-theme-canvas transition">
                  <span class="material-symbols-outlined text-[16px]">web</span> Default
                </button>
              </li>
              <li>
                <button onclick="changeThemePalette('fintech')" class="flex items-center gap-2 w-full text-left px-4 py-2 hover:bg-theme-canvas transition">
                  <span class="material-symbols-outlined text-[16px]">payments</span> Fintech
                </button>
              </li>
              <li>
                <button onclick="changeThemePalette('hig')" class="flex items-center gap-2 w-full text-left px-4 py-2 hover:bg-theme-canvas transition">
                  <span class="material-symbols-outlined text-[16px]">laptop_mac</span> macOS
                </button>
              </li>
            </ul>
          </div>'''

if target in content:
    content = content.replace(target, replacement)
    with open('Index.html', 'w') as f:
        f.write(content)
    print("Dropdown updated!")
else:
    print("Dropdown not found.")
