# Comvert Comments Slide Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Google Slides Apps Script that exports presentation comments and replies into a newly appended summary slide grouped by source slide.

**Architecture:** Keep comment-to-slide matching and export rendering in a pure Node-tested helper module, and keep Slides/Drive integration in a thin `Code.gs` layer. Map comments to slides by normalized quoted-text matching, then write one appended export slide with intra-deck links on each slide-group heading.

**Tech Stack:** Google Apps Script, Google Slides `SlidesApp`, Drive comments REST API via `UrlFetchApp`, Node built-in `assert`

---

### Task 1: Scaffold The Dedicated Subproject

**Files:**
- Create: `comvert-comments-slide/README.md`
- Create: `comvert-comments-slide/docs/requirements.md`
- Create: `comvert-comments-slide/docs/design.md`
- Create: `comvert-comments-slide/package.json`
- Create: `comvert-comments-slide/appsscript.json`
- Modify: `README.md`

- [ ] **Step 1: Add the new subproject docs and metadata**
- [ ] **Step 2: Add the new subproject entry to the repo root index**

### Task 2: Lock Pure Logic With Failing Tests First

**Files:**
- Create: `comvert-comments-slide/tests/convert_comments_slide.test.js`
- Create: `comvert-comments-slide/src/convert_comments_slide.js`

- [ ] **Step 1: Write failing tests for unique-match, ambiguous-match, unmatched, reply sorting, and export ordering**
- [ ] **Step 2: Run `npm test` in `comvert-comments-slide` and verify the new tests fail for the expected missing behavior**
- [ ] **Step 3: Implement the minimal pure helpers to satisfy those tests**
- [ ] **Step 4: Run `npm test` again and verify it passes**

### Task 3: Add Apps Script Integration

**Files:**
- Create: `comvert-comments-slide/Code.gs`
- Modify: `comvert-comments-slide/appsscript.json`

- [ ] **Step 1: Add `onOpen` menu registration and export entrypoint**
- [ ] **Step 2: Fetch comments from the Drive comments REST API using the active presentation file ID**
- [ ] **Step 3: Extract slide text, build the export model, append a new export slide, and apply group-heading slide links**
- [ ] **Step 4: Surface summary and error states through Slides UI alerts**

### Task 4: Verify Integration Surface

**Files:**
- Modify: `comvert-comments-slide/tests/convert_comments_slide.test.js`

- [ ] **Step 1: Add a regression check that `Code.gs` still exposes the expected menu, entrypoint, and slide-link calls**
- [ ] **Step 2: Run `npm test` in `comvert-comments-slide` and verify the full suite passes**
- [ ] **Step 3: Re-read the subproject docs and root index for consistency with the implemented behavior**
