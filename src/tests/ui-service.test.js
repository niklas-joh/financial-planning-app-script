/**
 * Financial Planning Tools - UI Service Tests
 *
 * This file contains tests for the FinancialPlanner.UIService module.
 * It includes mocking SpreadsheetApp, Ui, and HtmlService interactions.
 */
(function() {
  // Alias for easier access
  const T = FinancialPlanner.Testing;
  const UIService = FinancialPlanner.UIService;

  // --- Mock Dependencies & Globals ---
  let lastToast = null;
  let lastAlert = null;
  let lastPrompt = null;
  let lastSidebar = null;
  let lastModal = null;
  let alertResponse = null; // Control mock alert response
  let promptResponse = null; // Control mock prompt response { button: Button, text: string }

  const mockUi = {
    Button: { OK: 'OK', CANCEL: 'CANCEL', CLOSE: 'CLOSE' }, // Enum simulation
    ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO' }, // Enum simulation
    alert: function(title, message, buttons) {
      lastAlert = { title: title, message: message, buttons: buttons };
      return alertResponse || this.Button.OK; // Default response
    },
    prompt: function(title, message, buttons) {
       lastPrompt = { title: title, message: message, buttons: buttons };
       return {
           getSelectedButton: function() { return promptResponse ? promptResponse.button : mockUi.Button.CANCEL; },
           getResponseText: function() { return promptResponse ? promptResponse.text : null; }
       };
    },
    showSidebar: function(htmlOutput) {
        lastSidebar = { title: htmlOutput._title, content: htmlOutput._content, width: htmlOutput._width };
    },
    showModalDialog: function(htmlOutput, title) {
         lastModal = { title: title, content: htmlOutput._content, width: htmlOutput._width, height: htmlOutput._height };
    }
  };

  const mockHtmlOutput = {
      _title: null,
      _content: null,
      _width: 0,
      _height: 0,
      setTitle: function(title) { this._title = title; return this; },
      setWidth: function(width) { this._width = width; return this; },
      setHeight: function(height) { this._height = height; return this; },
      // Add other methods if needed by UIService
  };

  // Global mocks
  global.SpreadsheetApp = {
    getActiveSpreadsheet: function() {
      return {
        toast: function(message, title, timeoutSeconds) {
          lastToast = { message: message, title: title, timeout: timeoutSeconds };
        }
      };
    },
    getUi: function() { return mockUi; }
  };

  global.HtmlService = {
      createHtmlOutput: function(content) {
          // Return a new mock instance each time
          return Object.assign({}, mockHtmlOutput, { _content: content });
      }
  };
  
  global.Charts = { // Mock if Charts enum is used (e.g., in future chart service tests)
      ChartType: { PIE: 'PIE', COLUMN: 'COLUMN' } 
  };
  
  global.Session = { // Mock if Session is used (e.g., for time zone)
      getScriptTimeZone: function() { return "Etc/GMT"; } 
  };
  
  global.Utilities = { // Mock if Utilities is used (e.g., formatDate)
      formatDate: function(date, tz, format) { return `Formatted:${date.toISOString()}`; }
  };
  
  global.CacheService = { // Mock if CacheService is used (e.g., in dropdown tests)
      getScriptCache: function() { 
          return { 
              get: function(key){ return null; }, 
              put: function(key, value, ttl){},
              remove: function(key) {},
              removeAll: function(keys) {}
          }; 
      }
  };


  // --- Helper to reset mocks before each test ---
  function resetMocks() {
    lastToast = null;
    lastAlert = null;
    lastPrompt = null;
    lastSidebar = null;
    lastModal = null;
    alertResponse = mockUi.Button.OK; // Default OK
    promptResponse = { button: mockUi.Button.CANCEL, text: null }; // Default Cancel
  }

  // --- Test Cases ---

  T.registerTest("UIService", "showLoadingSpinner should call toast", function() {
    resetMocks();
    const message = "Loading data...";
    UIService.showLoadingSpinner(message);
    T.assertNotNull(lastToast, "Spreadsheet.toast should have been called.");
    T.assertEquals(message, lastToast.message, "Toast message should match.");
    T.assertEquals("Working...", lastToast.title, "Toast title should be 'Working...'.");
  });

  T.registerTest("UIService", "hideLoadingSpinner should call toast with empty message and short timeout", function() {
    resetMocks();
    UIService.hideLoadingSpinner();
    T.assertNotNull(lastToast, "Spreadsheet.toast should have been called for hiding.");
    T.assertEquals("", lastToast.message, "Toast message for hiding should be empty.");
    T.assertEquals("", lastToast.title, "Toast title for hiding should be empty.");
    T.assertEquals(1, lastToast.timeout, "Toast timeout for hiding should be 1.");
  });

  T.registerTest("UIService", "showSuccessNotification should call toast with 'Success' title", function() {
    resetMocks();
    const message = "Operation successful!";
    const duration = 10;
    UIService.showSuccessNotification(message, duration);
    T.assertNotNull(lastToast, "Spreadsheet.toast should have been called for success.");
    T.assertEquals(message, lastToast.message, "Success toast message should match.");
    T.assertEquals("Success", lastToast.title, "Success toast title should be 'Success'.");
    T.assertEquals(duration, lastToast.timeout, "Success toast duration should match.");
  });
  
   T.registerTest("UIService", "showSuccessNotification should use default duration", function() {
    resetMocks();
    const message = "Default duration test";
    UIService.showSuccessNotification(message); // No duration passed
    T.assertNotNull(lastToast, "Spreadsheet.toast should have been called for success.");
    T.assertEquals(message, lastToast.message, "Success toast message should match.");
    T.assertEquals("Success", lastToast.title, "Success toast title should be 'Success'.");
    T.assertEquals(5, lastToast.timeout, "Success toast duration should default to 5.");
  });

  T.registerTest("UIService", "showErrorNotification should call ui.alert", function() {
    resetMocks();
    const title = "Error Title";
    const message = "Detailed error message.";
    UIService.showErrorNotification(title, message);
    T.assertNotNull(lastAlert, "ui.alert should have been called for error.");
    // Note: The service formats the message with the title
    T.assertEquals(`${title}: ${message}`, lastAlert.message, "Alert message should include title and message.");
    // T.assertEquals(title, lastAlert.title, "Alert title should match."); // ui.alert combines title and message
  });

  T.registerTest("UIService", "showInfoAlert should call ui.alert with OK button", function() {
    resetMocks();
    const title = "Information";
    const message = "This is an informational message.";
    UIService.showInfoAlert(title, message);
    T.assertNotNull(lastAlert, "ui.alert should have been called for info.");
    T.assertEquals(title, lastAlert.title, "Info alert title should match.");
    T.assertEquals(message, lastAlert.message, "Info alert message should match.");
    T.assertEquals(mockUi.ButtonSet.OK, lastAlert.buttons, "Info alert should use OK button set.");
  });

  T.registerTest("UIService", "showConfirmationDialog should call ui.alert with OK_CANCEL and return true on OK", function() {
    resetMocks();
    alertResponse = mockUi.Button.OK; // Simulate user clicking OK
    const title = "Confirm Action";
    const message = "Are you sure?";
    const result = UIService.showConfirmationDialog(title, message);
    T.assertNotNull(lastAlert, "ui.alert should have been called for confirmation.");
    T.assertEquals(title, lastAlert.title, "Confirmation title should match.");
    T.assertEquals(message, lastAlert.message, "Confirmation message should match.");
    T.assertEquals(mockUi.ButtonSet.OK_CANCEL, lastAlert.buttons, "Confirmation should use OK_CANCEL button set.");
    T.assertTrue(result, "Should return true when user clicks OK.");
  });

  T.registerTest("UIService", "showConfirmationDialog should return false on Cancel", function() {
    resetMocks();
    alertResponse = mockUi.Button.CANCEL; // Simulate user clicking Cancel
    const result = UIService.showConfirmationDialog("Confirm", "Sure?");
    T.assertFalse(result, "Should return false when user clicks Cancel.");
  });

  T.registerTest("UIService", "showPromptDialog should call ui.prompt and return text on OK", function() {
    resetMocks();
    const inputText = "User input text";
    promptResponse = { button: mockUi.Button.OK, text: inputText }; // Simulate OK + text
    const title = "Enter Value";
    const message = "Please enter something:";
    const result = UIService.showPromptDialog(title, message);

    T.assertNotNull(lastPrompt, "ui.prompt should have been called.");
    T.assertEquals(title, lastPrompt.title, "Prompt title should match.");
    T.assertEquals(message, lastPrompt.message, "Prompt message should match.");
    T.assertEquals(mockUi.ButtonSet.OK_CANCEL, lastPrompt.buttons, "Prompt should use OK_CANCEL button set.");
    T.assertEquals(inputText, result, "Should return user text when OK is clicked.");
  });
  
   T.registerTest("UIService", "showPromptDialog should use defaultValue (though mock doesn't fully support it)", function() {
    resetMocks();
    // Note: The mock ui.prompt doesn't actually use the defaultValue, but we test the service calls it.
    // A more complex mock could simulate pre-filling.
    const defaultValue = "Default Text";
     promptResponse = { button: mockUi.Button.OK, text: defaultValue }; // Simulate OK returning default
    UIService.showPromptDialog("Title", "Message", defaultValue);
    // We mainly check the call was made correctly; the mock limitation prevents checking pre-fill.
    T.assertNotNull(lastPrompt, "ui.prompt should have been called.");
  });

  T.registerTest("UIService", "showPromptDialog should return null on Cancel", function() {
    resetMocks();
    promptResponse = { button: mockUi.Button.CANCEL, text: null }; // Simulate Cancel
    const result = UIService.showPromptDialog("Title", "Message");
    T.assertTrue(result === null, "Should return null when user clicks Cancel.");
  });
  
  T.registerTest("UIService", "showSidebar should call ui.showSidebar with configured HtmlOutput", function() {
      resetMocks();
      const title = "My Sidebar";
      const content = "<h1>Sidebar Content</h1>";
      UIService.showSidebar(title, content);
      
      T.assertNotNull(lastSidebar, "ui.showSidebar should have been called.");
      T.assertEquals(title, lastSidebar.title, "Sidebar title should match.");
      T.assertEquals(content, lastSidebar.content, "Sidebar content should match.");
      T.assertEquals(300, lastSidebar.width, "Sidebar width should default to 300."); // Default width check
  });
  
  T.registerTest("UIService", "showModalDialog should call ui.showModalDialog with configured HtmlOutput", function() {
      resetMocks();
      const title = "My Modal";
      const content = "<h2>Modal Content</h2>";
      const width = 500;
      const height = 350;
      UIService.showModalDialog(title, content, width, height);
      
      T.assertNotNull(lastModal, "ui.showModalDialog should have been called.");
      T.assertEquals(title, lastModal.title, "Modal title should match.");
      T.assertEquals(content, lastModal.content, "Modal content should match.");
      T.assertEquals(width, lastModal.width, "Modal width should match.");
      T.assertEquals(height, lastModal.height, "Modal height should match.");
  });
  
   T.registerTest("UIService", "showModalDialog should use default dimensions", function() {
      resetMocks();
      const title = "Default Modal";
      const content = "<p>Default dimensions</p>";
      UIService.showModalDialog(title, content); // Use defaults
      
      T.assertNotNull(lastModal, "ui.showModalDialog should have been called.");
      T.assertEquals(title, lastModal.title, "Modal title should match.");
      T.assertEquals(content, lastModal.content, "Modal content should match.");
      T.assertEquals(600, lastModal.width, "Modal width should default to 600.");
      T.assertEquals(400, lastModal.height, "Modal height should default to 400.");
  });


})(); // End IIFE
