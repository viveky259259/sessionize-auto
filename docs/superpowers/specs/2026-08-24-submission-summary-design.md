# Submission Summary Tool Design

## Goal

Add a local `create_submission_summary` MCP tool that produces an audience-directed narrative summary of conference submissions.

The tool must work from an existing prepared Session Reviewer workspace. When the caller supplies only an Excel workbook, it must first create a review workspace, assign topic categories, and run the existing quality and duplicate analysis before drafting the summary.

## Interface

The tool accepts:

- `workspace_dir` (optional): an existing Session Reviewer workspace.
- `source_path` (optional): an `.xlsx` workbook to import when no prepared workspace is supplied.
- `sheet_name` and `mapping` (optional): existing import controls, used only with `source_path`.
- `instruction` (required): free-form direction for the intended recipient, tone, emphasis, and desired length.

At least one of `workspace_dir` or `source_path` is required. A supplied workspace takes precedence. A supplied source workbook is imported into the specified workspace or a deterministic sibling review workspace.

## Workflow

1. Resolve the prepared workspace, importing the supplied workbook only when needed.
2. Ensure every session has a topic category and current quality/duplicate analysis.
3. Compute a compact factual summary: submission count, topic distribution, review progress, decision counts when present, quality-flag counts, and duplicate-cluster count.
4. Select representative highlights from titles, topics, and existing decisions without inventing claims about speakers or proposal content.
5. Return a structured response with the factual metrics, generated narrative, and the instruction used.

## Output

The result contains:

- `workspace_dir`
- `prepared_from_source` indicating whether the tool prepared the workspace
- `metrics` with count-based facts
- `topic_distribution`
- `summary` containing the recipient-directed draft

The narrative adapts only its framing and ordering to the instruction. It does not fabricate facts, recommendations, credentials, or Sessionize requirements.

## Safety and failure behavior

- The tool reads and writes only the local review workspace.
- It does not call Sessionize, submit ratings, publish comments, or alter the source workbook.
- A Sessionize organizer URL remains a skill-level, read-only preparation workflow; the MCP tool summarizes the local data produced by that workflow.
- The tool rejects a missing workspace/source pair, an invalid workbook, or a blank instruction with actionable errors.

## Tests

Cover:

1. Summary creation from a prepared workspace, including topic and duplicate facts.
2. Summary creation from a workbook that requires preparation.
3. Decision counts when reviews exist and omission-safe behavior when they do not.
4. Rejection of missing source/workspace and blank instructions.
5. A summary that preserves missing fields as unknown rather than inventing content.
