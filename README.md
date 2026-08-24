# Session Reviewer

Session Reviewer is a local Codex plugin for organizing and reviewing conference-session proposals. It provides an MCP server and a `session-review` skill for a fast, evidence-based review workflow.

## What it does

- Imports session proposals from an Excel workbook.
- Captures submitted Sessionize details through Codex's signed-in browser workflow, without changing Sessionize.
- Categorizes proposals by topic.
- Flags incomplete proposals and conservative duplicate clusters.
- Records five rubric scores, an accept/maybe/reject decision, and reviewer comments.
- Creates audience-directed factual summaries from prepared reviews or an incoming workbook.
- Exports reviewed sessions, a shortlist, and duplicate-cluster evidence to Excel.

## Privacy

The MCP server stores proposal data only in the workspace chosen for a review. It does not send proposal text to an external service. Sessionize access is read-only unless a user explicitly asks Codex to make a change.

Do not commit review workspaces or exported workbooks: they can contain submitted proposal data. The provided `.gitignore` excludes the common local artifacts.

## Using the plugin

Install the plugin in Codex from a marketplace that points to this source, then begin a new thread and ask Codex to review a session workbook or a signed-in Sessionize organizer URL. Codex will load the `session-review` skill and the `session-reviewer` MCP tools.

The core MCP flow is:

1. `import_sessions`
2. `enrich_session_details` (when Sessionize details were collected)
3. `categorize_topics`
4. `analyze_sessions`
5. `create_submission_summary` when a stakeholder-facing overview is needed
6. `get_review_batch` and `record_reviews`
7. `review_progress` and `export_reviews`

## Development

Requires Python 3.10 or newer and uses only the standard library.

```bash
python3 -m unittest discover -s server/tests -v
```

Validate the plugin and skill with the Codex plugin and skill validators before release.

## License

MIT. See [LICENSE](LICENSE).
