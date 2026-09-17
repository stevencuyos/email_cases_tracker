import re

with open('Index.html', 'r') as f:
    content = f.read()

# Locate the bottom-left dark mode toggle in the sidebar account section
target = r'''<button onclick="toggleDarkMode\(\)" class="text-gray-500 dark:text-gray-400 hover:bg-gray-200 dark:hover:bg-gray-700 p-1\.5 rounded-full transition" title="Toggle Dark Mode">
              <span id="themeIcon" class="material-symbols-outlined text-\[20px\]">light_mode</span>
            </button>'''

replacement = r'''<!-- Theme / Palette Dropdown -->
            <div class="relative inline-block text-left" id="themeDropdownContainer">
              <button onclick="toggleThemeMenu()" class="text-gray-500 dark:text-gray-400 hover:bg-gray-200 dark:hover:bg-gray-700 p-1.5 rounded-full transition" title="Change Theme">
                <span class="material-symbols-outlined text-[20px]">palette</span>
              </button>

              <!-- Dropdown Menu -->
              <div id="themeMenu" class="hidden absolute bottom-full left-0 mb-2 w-36 rounded-theme shadow-lg bg-theme-sidebar border border-theme-card-border overflow-hidden z-50">
                <div class="py-1">
                  <button onclick="changeTheme('default')" class="flex items-center gap-2 w-full text-left px-4 py-2 text-sm text-gray-700 dark:text-gray-200 hover:bg-theme-canvas transition">
                    <span class="material-symbols-outlined text-[16px]">web</span> Default
                  </button>
                  <button onclick="changeTheme('fintech')" class="flex items-center gap-2 w-full text-left px-4 py-2 text-sm text-gray-700 dark:text-gray-200 hover:bg-theme-canvas transition">
                    <span class="material-symbols-outlined text-[16px]">payments</span> Fintech
                  </button>
                  <button onclick="changeTheme('hig')" class="flex items-center gap-2 w-full text-left px-4 py-2 text-sm text-gray-700 dark:text-gray-200 hover:bg-theme-canvas transition">
                    <span class="material-symbols-outlined text-[16px]">laptop_mac</span> macOS
                  </button>
                </div>
              </div>
            </div>

            <button onclick="toggleDarkMode()" class="text-gray-500 dark:text-gray-400 hover:bg-gray-200 dark:hover:bg-gray-700 p-1.5 rounded-full transition" title="Toggle Dark Mode">
              <span id="themeIcon" class="material-symbols-outlined text-[20px]">light_mode</span>
            </button>'''

if target in content:
    content = content.replace(target, replacement)
    with open('Index.html', 'w') as f:
        f.write(content)
    print("Toggle button replaced successfully.")
else:
    print("Target string not found for toggle replacement.")
