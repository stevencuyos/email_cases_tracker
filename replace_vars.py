import re

with open('Index.html', 'r') as f:
    content = f.read()

# Let's target the remaining bg-gray-50 which are used for hovers and table headers,
# replacing them with our new variables so that themes like HIG or Fintech can look consistent.
# The user said "replace all hardcoded ones".

# We'll use bg-gray-50/50 -> bg-theme-canvas/50
# We did this mostly, but let's do:
# hover:bg-gray-50 -> hover:bg-gray-100 (if it's a utility hover on a white card) or hover:bg-theme-canvas if that makes sense.
# Actually, the user asked to "replace all hardcoded background and border-radius classes".
# Let's replace:
# bg-gray-50 -> bg-theme-canvas
# But we already did that except for the dark: and / variations. Let's do a strict replacement.

content = re.sub(r'\bbg-gray-50\b', 'bg-theme-canvas', content)
content = re.sub(r'\bbg-gray-50/(\d+)\b', r'bg-theme-canvas/\1', content)

# bg-white inside classes might need to be bg-theme-sidebar or bg-theme-card depending on usage.
# But it's risky to replace ALL bg-white blindly. Let's stick to what's obviously layout.
# The user specified:
# --color-canvas, --bg-card, --radius-card, --shadow-card.

with open('Index.html.new', 'w') as f:
    f.write(content)
