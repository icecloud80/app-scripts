# Comvert Comments Slide Design

## Goal

Build a Google Slides-bound Apps Script that exports presentation comments into a newly appended summary slide at the end of the deck.

## Scope

- Read all top-level comments and replies for the active presentation.
- Group top-level comment threads by source slide.
- Sort top-level threads by creation time within each slide group.
- Append a new export slide on every run.
- Add one slide link on each slide-group heading.
- Include explicit fallback groups for comments that cannot be mapped to one slide.

## Non-Goals

- Do not mutate or delete existing export slides.
- Do not guarantee a perfect comment-to-slide mapping when the source text is duplicated or missing.
- Do not export resolved or deleted comments specially beyond skipping deleted entries returned by the API.
- Do not build a sidebar, modal, or multi-step UI.

## Constraints

- Google Apps Script does not expose a stable documented direct mapping from Drive comments to Google Slides page IDs.
- Drive comments can expose `quotedFileContent.value`, replies, timestamps, and author info through the Drive comments API.
- Slides Apps Script exposes slide enumeration, page element text access, and intra-deck links via `TextStyle.setLinkSlide(slide)`.

## Recommended Mapping Strategy

Use a best-effort text matching strategy:

1. Enumerate every normal slide in the active presentation.
2. Extract normalized visible text from shapes and table cells on each slide.
3. For each top-level comment thread:
   - Read `quotedFileContent.value`.
   - If the normalized quote matches exactly one slide text corpus, map the thread to that slide.
   - If it matches multiple slides, route the thread to `Ambiguous Comments`.
   - If the quote is empty or matches no slides, route the thread to `Unmatched Comments`.
4. Keep replies nested under their parent thread and sort replies by creation time.

This is preferred over parsing comment anchors because the Drive API does not document a stable Slides anchor format.

## Output Slide Layout

The script appends one new slide at the end on every run with:

- Title: `Comments Export`
- Subtitle: export timestamp, presentation title, and thread count summary
- Body content grouped in this order:
  - slide groups ordered by slide index
  - `Ambiguous Comments` if non-empty
  - `Unmatched Comments` if non-empty

Each slide group contains:

- A heading line like `Slide 3 - Roadmap Overview`
- The entire heading text linked to the source slide
- One block per top-level thread:
  - metadata line: `Person: X | Time: Y`
  - comment body line
  - zero or more indented reply blocks with the same metadata + body structure

## Project Structure

Keep this as a dedicated top-level Apps Script subproject at `comvert-comments-slide/` to match the requested boundary and existing repo convention.

- `comvert-comments-slide/Code.gs`
  - menu entrypoint
  - Drive comments API integration
  - slide extraction
  - export-slide writing
- `comvert-comments-slide/src/convert_comments_slide.js`
  - pure matching, sorting, grouping, and rendering-model helpers
- `comvert-comments-slide/tests/convert_comments_slide.test.js`
  - Node tests for pure logic and generated output model
- `comvert-comments-slide/README.md`
  - usage summary and limits
- `comvert-comments-slide/docs/requirements.md`
  - product requirements
- `comvert-comments-slide/docs/design.md`
  - subproject design

## Error Handling

- If the active file is not a Slides presentation, surface a UI alert and stop.
- If Drive comments fetch fails, surface the HTTP status and message.
- If no comments are returned, still append an export slide with a `No comments found.` body.
- If no slide can be written due to shape insertion issues, fail loudly rather than pretending success.

## Verification Strategy

- Test-first coverage for:
  - unique slide match
  - ambiguous multi-slide match
  - unmatched comments
  - reply sorting and nesting
  - slide-group ordering
  - no-comments export model
- After pure logic passes, verify the Apps Script file still exposes:
  - menu registration
  - export entrypoint
  - slide-link application

