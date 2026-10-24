// ESLint config for a Google Apps Script project.
//
// Apps Script concatenates every file into one shared global scope, so a
// function declared in one file is a global in all the others. ESLint only
// sees one file at a time, so we collect the top-level names from src/ and
// declare them as globals; `no-undef` then catches genuine typos instead of
// flagging every cross-file call.
import js from '@eslint/js';
import globals from 'globals';
import { collectTopLevelNames } from './scripts/lib/source-files.mjs';

const appsScriptGlobals = Object.fromEntries(
  [
    'CacheService',
    'CalendarApp',
    'Charts',
    'ContentService',
    'DriveApp',
    'GmailApp',
    'HtmlService',
    'LockService',
    'Logger',
    'MailApp',
    'PropertiesService',
    'ScriptApp',
    'console',
    'Session',
    'SpreadsheetApp',
    'UrlFetchApp',
    'Utilities',
  ].map((name) => [name, 'readonly']),
);

const projectGlobals = Object.fromEntries(
  collectTopLevelNames().map((name) => [name, 'writable']),
);

export default [
  { ignores: ['node_modules/**', '.artifacts/**'] },
  js.configs.recommended,
  {
    files: ['src/**/*.js'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'script',
      globals: { ...appsScriptGlobals, ...projectGlobals },
    },
    linterOptions: {
      // Older files carry disable comments for rules from a previous config.
      reportUnusedDisableDirectives: 'off',
    },
    rules: {
      // Top-level functions are called by Apps Script (menus, triggers,
      // google.script.run), so they look unused to ESLint.
      'no-unused-vars': ['error', { vars: 'local', args: 'none', caughtErrors: 'none' }],
      // `var FinancialPlanner = FinancialPlanner || {}` in every file is the
      // intended pattern for the shared namespace.
      'no-redeclare': ['error', { builtinGlobals: false }],
      eqeqeq: ['error', 'smart'],
      'no-implicit-globals': 'off',
    },
  },
  {
    files: ['src/tests/**/*.js'],
    languageOptions: {
      // The tests stub Apps Script services on Node's `global`.
      globals: { global: 'writable' },
    },
    rules: {
      'preserve-caught-error': 'off',
      // Test files register closures and stub objects freely.
      'no-unused-vars': 'off',
    },
  },
  {
    // bank-sync is a separate Apps Script project. Its files also export via
    // `module.exports` when loaded in Node, so its tests can require them.
    files: ['bank-sync/src/**/*.js'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'script',
      globals: { ...appsScriptGlobals, BankSync: 'writable', module: 'readonly' },
    },
    rules: {
      // Entry points (onOpen, doGet, menu handlers) are called by Apps Script.
      'no-unused-vars': ['error', { vars: 'local', args: 'none', caughtErrors: 'none' }],
    },
  },
  {
    files: ['bank-sync/test/**/*.js'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'commonjs',
      globals: globals.node,
    },
  },
  {
    files: ['scripts/**/*.mjs', 'eslint.config.js'],
    languageOptions: {
      ecmaVersion: 2024,
      sourceType: 'module',
      globals: globals.node,
    },
  },
];
