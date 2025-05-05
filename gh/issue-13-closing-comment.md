## Implementation Complete

I've implemented the ability to toggle sub-category view in the overview sheet as requested. The implementation includes:

1. **Menu Option**: Added a new toggle option under Settings menu to show/hide sub-categories
2. **In-Sheet Toggle**: Added a checkbox directly in the Overview sheet for quick toggling
3. **Persistence**: User preference is stored and remembered between sessions

### Technical Details

- Added user preference system in the settings.js file
- Modified the overview generation logic to respect the sub-category visibility setting
- Added event handlers to detect checkbox changes and update the view accordingly
- Implemented visual feedback when toggling the setting

This implementation allows users to choose between a detailed view with all sub-categories (current behavior) or a simplified view with only categories shown (aggregating all sub-categories).

Closes #13
