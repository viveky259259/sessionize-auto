---
name: session-review
description: Build or review a conference session queue from an Excel workbook or Sessionize organizer view using local Session Reviewer MCP tools. Use for description capture, topic categorization, duplicate comparison, scoring, and decision export.
---

# Session Review

Use the local `session-reviewer` MCP tools to run an evidence-based review without sending workbook content to an external service.

## Sessionize preparation

When the user supplies a Sessionize organizer URL, use the signed-in Sessionize view only to read submitted proposal data. First look for an applicable Sessionize connector; if none is available, use the browser workflow. Do not submit ratings, comments, or other changes in Sessionize unless the user explicitly asks.

1. Reuse an existing review workbook when present; otherwise create one in the chosen output location.
2. Collect the session ID, title, speaker, format, level, status, detail URL, and submitted Description for every proposal. Persist collected data in bounded batches so a slow page does not lose earlier work.
3. Call `enrich_session_details` with the collected details, then `categorize_topics`. Keep a missing submitted description blank.
4. Update the review workbook with Description immediately after Title, preserve its existing review fields, and add the topic category field. Make descriptions wrapped and readable.
5. Use `analyze_sessions` after enrichment, because descriptions improve duplicate and quality analysis.

## Submission summaries

Use `create_submission_summary` when the user asks for an overview intended for another person. Provide their direction as `instruction`. The tool reuses a prepared workspace when available; otherwise give it the source workbook so it can prepare local categories and analysis first. Treat the returned narrative as a factual draft and keep any final audience-specific claims rooted in its metrics and titles.

## Review flow

1. Call `import_sessions` for the supplied `.xlsx` file. If field inference is ambiguous, call `inspect_import`, explain the alternatives, and ask the user to choose a mapping before reviewing.
2. Call `analyze_sessions`. Explain that quality flags and duplicate clusters are cues, not automatic rejections.
3. Fetch unreviewed proposals in batches of 10–15 with `get_review_batch`. Score each proposal before moving to the next batch. Save every completed batch through `record_reviews` so progress persists.
4. Use `get_duplicate_cluster` to compare similar proposals. Recommend the strongest session in the cluster, but do not reject a proposal merely for being similar if its angle, target audience, or speaker clearly differs.
5. Use `review_progress` to keep the user informed. Call `export_reviews` after completion or when the user requests an interim result.

## Rubric

Score each dimension from 1 to 5:

- **Clarity and depth:** Is the proposal specific, well structured, and substantive?
- **Audience fit:** Does it suit the stated event audience and track?
- **Speaker credibility:** Does the provided speaker information support the promised content? Do not infer credentials that are not supplied.
- **Originality:** Does it bring a distinct perspective compared with similar submissions?
- **Delivery readiness:** Is the scope, format, and expected takeaway ready to program?

Set the decision to `accept`, `maybe`, or `reject`. Give a concise, constructive comment for every `maybe` and `reject`, and when a weak score materially influences an `accept` decision. Keep comments rooted in the proposal's actual text.

## Guardrails

- Treat missing information as unknown, not negative evidence.
- Never fabricate speaker backgrounds, event requirements, or content claims.
- Do not expose proposal text outside the local review workspace.
- Do not overwrite a source workbook. `export_reviews` creates a separate file and requires an explicit overwrite flag for an existing target.
- Before claiming the review is complete, use `review_progress` and report the reviewed count and unresolved duplicate clusters.
