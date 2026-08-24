import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

MODULE = Path(__file__).parents[1] / "session_reviewer_mcp.py"
SPEC = importlib.util.spec_from_file_location("session_reviewer_mcp", MODULE)
SERVER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(SERVER)


class SessionReviewerTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.source = self.root / "submissions.xlsx"
        SERVER.write_xlsx(self.source, [("Sessions", ["Session ID", "Title", "Description", "Speaker", "Track"], [["s-1", "Secure APIs", "A detailed practical guide to securing application programming interfaces with real-world examples and a threat model.", "Ada", "Security"], ["s-2", "Secure API", "A detailed guide to securing APIs with practical examples, threat models, and implementation advice.", "Ben", "Security"], ["s-3", "Lightning Intro", "Short.", "", "General"]])])
        self.workspace = self.root / "review-state"

    def tearDown(self):
        self.temp.cleanup()

    def test_import_analysis_review_and_export(self):
        imported = SERVER.tool_import({"source_path": str(self.source), "workspace_dir": str(self.workspace)})
        self.assertEqual(imported["session_count"], 3)
        self.assertEqual(imported["mapping"]["title"], "Title")
        enriched = SERVER.tool_enrich_details({"workspace_dir": str(self.workspace), "sessions": [{"id": "s-1", "description": "A hands-on Flutter session for building secure APIs.", "sessionize_url": "https://sessionize.com/app/organizer/session/1/s-1", "format": "Session", "level": "Intermediate", "status": "Nominated"}]})
        self.assertEqual(enriched["updated_count"], 1)
        categorized = SERVER.tool_categorize_topics({"workspace_dir": str(self.workspace)})
        self.assertEqual(categorized["categories"]["Security & Privacy"], 2)
        analysis = SERVER.tool_analyze({"workspace_dir": str(self.workspace)})
        self.assertEqual(analysis["duplicate_cluster_count"], 1)
        batch = SERVER.tool_batch({"workspace_dir": str(self.workspace), "limit": 3})
        self.assertIn("short_abstract", batch["sessions"][2]["quality_flags"])
        saved = SERVER.tool_record({"workspace_dir": str(self.workspace), "reviews": [{"id": "s-1", "scores": {"clarity_depth": 5, "audience_fit": 4, "speaker_credibility": 4, "originality": 4, "delivery_readiness": 5}, "decision": "accept", "comment": "Specific and useful."}, {"id": "s-2", "scores": {"clarity_depth": 4, "audience_fit": 4, "speaker_credibility": 3, "originality": 3, "delivery_readiness": 4}, "decision": "maybe", "comment": "Compare with s-1."}, {"id": "s-3", "scores": {"clarity_depth": 1, "audience_fit": 2, "speaker_credibility": 1, "originality": 2, "delivery_readiness": 1}, "decision": "reject", "comment": "Needs a substantive abstract."}]})
        self.assertEqual(saved["reviewed_total"], 3)
        output = self.root / "reviewed.xlsx"
        exported = SERVER.tool_export({"workspace_dir": str(self.workspace), "output_path": str(output)})
        self.assertFalse(exported["partial_review"])
        self.assertTrue(zipfile_names(output, "xl/worksheets/sheet3.xml"))
        _, headers, rows = SERVER.read_xlsx(output, "Reviewed Sessions")
        self.assertIn("Decision", headers)
        self.assertIn("Topic Category", headers)
        self.assertEqual(rows[0]["Topic Category"], "Security & Privacy")
        self.assertEqual(rows[0]["Decision"], "accept")

        summary = SERVER.tool_create_summary({"workspace_dir": str(self.workspace), "instruction": "A concise organizer update."})
        self.assertFalse(summary["prepared_from_source"])
        self.assertEqual(summary["metrics"]["submission_count"], 3)
        self.assertEqual(summary["metrics"]["decision_counts"], {"accept": 1, "maybe": 1, "reject": 1})
        self.assertIn("Prepared for: A concise organizer update.", summary["summary"])

    def test_invalid_score_is_rejected_without_state_mutation(self):
        SERVER.tool_import({"source_path": str(self.source), "workspace_dir": str(self.workspace)})
        with self.assertRaises(SERVER.ToolError):
            SERVER.tool_record({"workspace_dir": str(self.workspace), "reviews": [{"id": "s-1", "scores": {"clarity_depth": 6, "audience_fit": 4, "speaker_credibility": 4, "originality": 4, "delivery_readiness": 4}, "decision": "accept"}]})
        state = json.loads((self.workspace / "session-review-state.json").read_text())
        self.assertEqual(state["reviews"], {})

    def test_rejects_unknown_detail_fields(self):
        SERVER.tool_import({"source_path": str(self.source), "workspace_dir": str(self.workspace)})
        with self.assertRaises(SERVER.ToolError):
            SERVER.tool_enrich_details({"workspace_dir": str(self.workspace), "sessions": [{"id": "s-1", "unknown": "not allowed"}]})

    def test_explicit_blank_description_does_not_fall_back_to_source_abstract(self):
        SERVER.tool_import({"source_path": str(self.source), "workspace_dir": str(self.workspace)})
        SERVER.tool_enrich_details({"workspace_dir": str(self.workspace), "sessions": [{"id": "s-1", "description": ""}]})
        batch = SERVER.tool_batch({"workspace_dir": str(self.workspace), "limit": 1})
        self.assertEqual(batch["sessions"][0]["abstract"], "")

    def test_malformed_request_does_not_stop_following_valid_mcp_request(self):
        process = subprocess.run(
            [sys.executable, str(MODULE)],
            input='not-json\n{"jsonrpc":"2.0","id":1,"method":"initialize","params":{}}\n',
            text=True,
            capture_output=True,
            check=True,
        )
        responses = [json.loads(line) for line in process.stdout.splitlines()]
        self.assertEqual(len(responses), 1)
        self.assertEqual(responses[0]["id"], 1)

    def test_summary_prepares_source_workbook_and_rejects_missing_inputs(self):
        workspace = self.root / "summary-workspace"
        summary = SERVER.tool_create_summary({"source_path": str(self.source), "workspace_dir": str(workspace), "instruction": "A neutral intake overview."})
        self.assertTrue(summary["prepared_from_source"])
        self.assertEqual(summary["metrics"]["submission_count"], 3)
        self.assertIn("Security & Privacy", summary["topic_distribution"])
        self.assertTrue((workspace / "session-review-state.json").is_file())
        with self.assertRaises(SERVER.ToolError):
            SERVER.tool_create_summary({"instruction": ""})


def zipfile_names(path, expected):
    import zipfile
    with zipfile.ZipFile(path) as archive:
        return expected in archive.namelist()


if __name__ == "__main__":
    unittest.main()
