# Development Principles

## Architecture Overview

This Google Apps Script project follows a simplified, execution-order-independent architecture that prioritizes maintainability and robustness.

## Core Principles

### 1. KISS (Keep It Simple, Stupid)

**What we do:**
- Use direct object literal assignments instead of complex constructor patterns
- Avoid unnecessary abstraction layers
- Write straightforward, readable code

**What we avoid:**
- Over-engineered patterns that add complexity without value
- Multiple levels of indirection
- Clever code that's hard to understand

### 2. YAGNI (You Aren't Gonna Need It)

**What we do:**
- Implement features only when actually needed
- Use the simplest solution that works
- Remove unused code and patterns

**What we avoid:**
- Premature optimization
- "Future-proofing" with unnecessary flexibility
- Building infrastructure before it's needed

### 3. Single Responsibility Principle (SRP)

**What we do:**
- Each module has one clear purpose
- Services are focused and cohesive
- Functions do one thing well

**What we avoid:**
- God objects that do everything
- Mixing unrelated concerns
- Functions with multiple responsibilities

## Module Pattern

### Object Literal Pattern (Recommended)

All services use the object literal pattern for simplicity:

```javascript
// Ensure namespace exists
var FinancialPlanner = FinancialPlanner || {};

// Direct assignment - no constructors needed
FinancialPlanner.ServiceName = {
  method1: function() {
    // Implementation
  },
  
  method2: function() {
    // Can access other services via namespace
    FinancialPlanner.OtherService.someMethod();
  }
};
```

### IIFE for Private State (When Needed)

When you need private variables or helper functions:

```javascript
FinancialPlanner.ServiceName = (function() {
  // Private variables (closure)
  const privateVar = 'private';
  
  // Private helper function
  function privateHelper() {
    return privateVar;
  }
  
  // Public API
  return {
    publicMethod: function() {
      return privateHelper();
    }
  };
})();
```

### What We DON'T Do

❌ **Constructor Pattern with Separate Instantiation**
```javascript
// AVOID THIS - Creates execution order dependencies
const ServiceModule = (function() {
  function ServiceConstructor(dependency1, dependency2) {
    this.dep1 = dependency1;
    this.dep2 = dependency2;
  }
  return ServiceConstructor;
})();

// Separate instantiation - FRAGILE!
FinancialPlanner.Service = new ServiceModule(dep1, dep2);
```

## File Execution Order

### Key Insight

**Google Apps Script executes files alphabetically, NOT in `filePushOrder`.**

The `filePushOrder` in `.clasp.json` controls deployment order, not execution order.

### Our Solution

We use **method-time dependencies** instead of **file-load-time dependencies**:

```javascript
// ✅ GOOD - Dependencies accessed when method runs
FinancialPlanner.ServiceA = {
  doSomething: function() {
    // These modules exist by the time this method is called
    const config = FinancialPlanner.Config.get();
    FinancialPlanner.ServiceB.helperMethod();
  }
};

// ❌ BAD - Dependencies accessed at file load time
const ServiceA = new ServiceConstructor(
  FinancialPlanner.Config,  // Might not exist yet!
  FinancialPlanner.ServiceB // Might not exist yet!
);
```

## Dependency Management

### Accessing Dependencies

Services access each other through the global `FinancialPlanner` namespace:

```javascript
FinancialPlanner.MyService = {
  myMethod: function() {
    // Access config
    const sheetNames = FinancialPlanner.Config.getSheetNames();
    
    // Call another service
    FinancialPlanner.ErrorService.handle(error, message);
    
    // Use utilities
    const colLetter = FinancialPlanner.Utils.columnToLetter(5);
  }
};
```

### No Dependency Injection Needed

Since all modules are singletons accessed via namespace, there's no need for dependency injection at construction time.

## Google Apps Script Specifics

### Namespace Pattern

Always ensure the namespace exists:

```javascript
// At the top of every module file
var FinancialPlanner = FinancialPlanner || {};
```

### Global Trigger Functions

Apps Script triggers (onOpen, onEdit) must be global functions:

```javascript
// Global function that delegates to namespace
function onOpen() {
  if (FinancialPlanner && FinancialPlanner.Controllers) {
    FinancialPlanner.Controllers.onOpen();
  }
}
```

### HTML Files

HTML files in the project root (e.g., `services/plaid-link.html`) are referenced without the `src/` prefix:

```javascript
// Correct - no src/ prefix
const html = HtmlService.createHtmlOutputFromFile('services/plaid-link');
```

## Best Practices

### 1. Prefer Simple Over Clever

```javascript
// ✅ GOOD - Clear and simple
function calculateTotal(items) {
  let total = 0;
  for (let i = 0; i < items.length; i++) {
    total += items[i].amount;
  }
  return total;
}

// ❌ AVOID - Clever but harder to read
const calculateTotal = items => items.reduce((sum, {amount}) => sum + amount, 0);
```

### 2. Use Descriptive Names

```javascript
// ✅ GOOD
function buildMonthlySumFormula(params) { }

// ❌ BAD
function bmsf(p) { }
```

### 3. Document Public APIs

Use JSDoc for all public methods:

```javascript
/**
 * Calculates the savings rate.
 * @param {number} income - Total income
 * @param {number} savings - Total savings
 * @returns {number} Savings rate as decimal
 * @memberof FinancialPlanner.MetricsCalculator
 */
calculateSavingsRate: function(income, savings) {
  if (income === 0) return 0;
  return savings / income;
}
```

### 4. Handle Errors Gracefully

```javascript
myMethod: function() {
  try {
    // Operation
  } catch (error) {
    FinancialPlanner.ErrorService.handle(error, 'User-friendly message');
    throw error; // Re-throw if needed
  }
}
```

### 5. Use Method-Time Resolution

```javascript
// ✅ GOOD - Resolves at method call time
getSheetName: function() {
  return FinancialPlanner.Config.getSheetNames().OVERVIEW;
}

// ❌ BAD - Tries to resolve at file load time
const SHEET_NAME = FinancialPlanner.Config.getSheetNames().OVERVIEW;
```

## Project Structure

```
src/
├── core/              # Core application logic
│   ├── config.js      # Configuration management
│   ├── controllers.js # UI action coordination
│   └── index.js       # Additional initialization
├── services/          # Reusable services
│   ├── ui-service.js
│   ├── error-service.js
│   ├── cache-service.js
│   └── ...
├── features/          # Feature-specific code
│   ├── financial-overview/
│   ├── financial-analysis/
│   └── ...
└── utils/             # Utility functions
    └── common.js
```

## Testing

### Manual Testing Checklist

After making changes:

1. ✅ Deploy with `clasp push`
2. ✅ Open the Google Sheet
3. ✅ Verify the menu appears
4. ✅ Test a simple menu action (e.g., "Generate Overview")
5. ✅ Check for errors in Execution log (Extensions > Apps Script > Executions)

### What to Watch For

- Undefined reference errors (service not loaded)
- Execution order issues (dependencies not ready)
- Menu not appearing (onOpen not firing)

## Clasp Configuration

### .clasp.json

```json
{
  "scriptId": "your-script-id",
  "rootDir": "src",
  "filePushOrder": [
    "utils/common.js",
    "core/config.js",
    "services/ui-service.js",
    // ... other files
    "core/controllers.js"
  ]
}
```

**Note:** `filePushOrder` affects deployment order, not execution order. It's mainly useful for:
- Keeping related files together in the Apps Script editor
- Ensuring HTML files and dependencies deploy in logical order

## Common Pitfalls

### ❌ Don't Use `new` with Services

```javascript
// WRONG
const config = new FinancialPlanner.Config();

// RIGHT
const value = FinancialPlanner.Config.getValue();
```

### ❌ Don't Create Multiple Instances

```javascript
// WRONG - Services are singletons
const myConfig = { ...FinancialPlanner.Config };

// RIGHT - Use the service directly
FinancialPlanner.Config.update({ ... });
```

### ❌ Don't Rely on File Load Order

```javascript
// WRONG - Might not exist yet
const sheetNames = FinancialPlanner.Config.getSheetNames();

// RIGHT - Access when needed
function myFunction() {
  const sheetNames = FinancialPlanner.Config.getSheetNames();
}
```

## Migration from Constructor Pattern

If you encounter old constructor-based code:

1. Remove the IIFE that returns a constructor
2. Convert constructor function to object literal
3. Move dependency resolution from constructor to methods
4. Update any instantiation code (delete `new` calls)

## Summary

**Keep it simple.** Use object literals. Access dependencies through the namespace when methods run, not when files load. This architecture is robust, maintainable, and aligns with Google Apps Script's execution model.
