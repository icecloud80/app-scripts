# Convert Comments Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Google Docs-bound Apps Script that exports comments from the current active tab into a plain-text feedback block.

**Architecture:** Keep matching and formatting logic in a testable Node module, and keep Google Docs / Drive integration in a thin `Code.gs` layer. Read comments through the Drive comments REST API, infer tab ownership via quoted text matching, and write output either into an existing `Feedback` tab or into the current tab as a fallback.

**Tech Stack:** Google Apps Script, Google Docs `DocumentTab`, Drive comments REST API via `UrlFetchApp`, Node built-in `assert`

---

### Task 1: Scaffold the Project

**Files:**
- Create: `convert-comments/README.md`
- Create: `convert-comments/docs/requirements.md`
- Create: `convert-comments/docs/design.md`
- Create: `convert-comments/package.json`
- Create: `convert-comments/appsscript.json`

- [ ] **Step 1: Write the project docs and metadata**
- [ ] **Step 2: Confirm the new project directory matches existing repo conventions**

### Task 2: Lock Pure Logic with Tests

**Files:**
- Create: `convert-comments/tests/convert_comments.test.js`
- Create: `convert-comments/src/convert_comments.js`

- [ ] **Step 1: Write a failing test for grouping comments by highlight text**
- [ ] **Step 2: Run `npm test` in `convert-comments` and verify it fails**
- [ ] **Step 3: Implement minimal grouping and formatting helpers**
- [ ] **Step 4: Run `npm test` again and verify it passes**

### Task 3: Add Apps Script Integration

**Files:**
- Create: `convert-comments/Code.gs`
- Modify: `convert-comments/appsscript.json`

- [ ] **Step 1: Add `onOpen` menu and conversion entrypoint**
- [ ] **Step 2: Fetch comments through Drive comments REST API**
- [ ] **Step 3: Detect `Feedback` tab or append fallback output**
- [ ] **Step 4: Surface summary and failure messages through Docs UI**

### Task 4: Update Repo Index and Verify

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Add the new subproject to the repo index**
- [ ] **Step 2: Run `npm test` in `convert-comments`**
- [ ] **Step 3: Review the generated files and behavior notes**
