# Issue #19 Fix: "Identifier 'CONFIG' has already been declared"

## Problem
When generating the financial overview, the following error occurred:
```
SyntaxError: Identifier 'CONFIG' has already been declared
```

## Root Cause
After investigation, I found that both `src/features/finance_overview.js` and `src/features/dropdowns.js` declare a constant named `CONFIG`. In Google Apps Script, all script files are concatenated and run in the same scope, causing a naming conflict when the same variable name is declared twice.

## Solution
1. Renamed the `CONFIG` constant in `finance_overview.js` to `OVERVIEW_CONFIG` to make it unique
2. Updated all references to `CONFIG` in `finance_overview.js` to use `OVERVIEW_CONFIG` instead
3. Left the `CONFIG` constant in `dropdowns.js` unchanged since it was declared first

## Changes Made
1. In `src/features/finance_overview.js`:
   - Renamed `const CONFIG = {...}` to `const OVERVIEW_CONFIG = {...}`
   - Updated all references to `CONFIG.SHEETS`, `CONFIG.TYPE_ORDER`, `CONFIG.EXPENSE_TYPES`, etc. to use `OVERVIEW_CONFIG` instead

## Verification
The changes should resolve the naming conflict because:
1. Each constant now has a unique name (`CONFIG` in dropdowns.js and `OVERVIEW_CONFIG` in finance_overview.js)
2. All references to the renamed constant have been updated
3. The functionality remains unchanged since we only modified the variable name

## Next Steps
1. Test the financial overview generation to confirm the error is resolved
2. Consider adding a naming convention for constants across the codebase to prevent similar issues in the future
