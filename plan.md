1.  **Refactor Color Variables:**
    *   Currently, colors are hardcoded as specific Tailwind colors (e.g., `bg-blue-600`, `bg-gray-50`, `bg-white`).
    *   Extract the current "Default" colors into CSS variables and set them in `.theme-default`.
    *   Create variables for `--color-canvas`, `--bg-card`, `--radius-card`, `--shadow-card`, `--color-primary`, `--color-accent`, etc.
    *   Update `tailwind.config` to use these CSS variables.
    *   Update HTML to use the new dynamic Tailwind classes.
2.  **Implement Fintech Theme:**
    *   Create `.theme-fintech` class with Stripe-like design variables.
    *   Characteristics: sharp/minimal border radiuses (`rounded-lg`), stark white backgrounds against a very light cool-gray canvas, sophisticated/muted accent colors, and subtle, diffuse drop shadows.
3.  **Implement HIG Theme:**
    *   Create `.theme-hig` class with Apple HIG (Glassmorphism) variables.
    *   Characteristics: heavy rounding (`rounded-3xl`), translucent panel backgrounds with backdrop blur (`bg-white/70 backdrop-blur-lg`), slightly thicker typography/icons, vibrant/saturated accent colors.
4.  **Implement UI Toggle:**
    *   Add a "Theme/Palette" icon button (using the `palette` material symbol) to the bottom-right corner, near the light/dark mode button.
    *   Ensure the dropdown allows selection of "Default", "Fintech", or "macOS".
5.  **Implement Theme Logic:**
    *   Write JavaScript to handle theme switching (`changeTheme(themeName)`).
    *   Save the selected theme in `localStorage`.
    *   Ensure the theme persists across reloads.
    *   Verify the charts DO NOT change themes.
    *   Ensure that each theme handles light/dark mode correctly via media queries or `.dark` class nesting within the theme class.
6.  **Pre-commit checks**
    *   Complete pre commit steps to make sure proper testing, verifications, reviews and reflections are done.
7.  **Submit changes**
    *   Once all steps are completed and verified, commit the changes.
