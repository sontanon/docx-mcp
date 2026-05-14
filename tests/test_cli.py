"""Tests for the CLI module."""

import json

from docx_mcp.cli import main


class TestCliHelp:
    def test_help_returns_zero(self, capsys):
        try:
            main(["--help"])
        except SystemExit as e:
            assert e.code == 0
        captured = capsys.readouterr()
        assert "apply" in captured.out
        assert "convert" in captured.out
        assert "validate" in captured.out


class TestCliConvert:
    def test_convert_tagged(self, simple_5para_path, capsys):
        rc = main(["convert", str(simple_5para_path)])
        assert rc == 0
        out = capsys.readouterr().out
        assert "<f=1>" in out
        assert "<f=5>" in out

    def test_convert_json(self, simple_5para_path, capsys):
        rc = main(["convert", str(simple_5para_path), "--format", "json"])
        assert rc == 0
        out = capsys.readouterr().out
        data = json.loads(out)
        assert len(data) == 5
        assert data[0]["fragment_id"] == 1
        assert "Seller" in data[0]["text"]

    def test_convert_missing_file(self, capsys):
        rc = main(["convert", "/nonexistent/file.docx"])
        assert rc == 1
        err = capsys.readouterr().err
        assert "not found" in err


class TestCliApply:
    def _write_changes(self, tmp_path, changes):
        p = tmp_path / "changes.json"
        p.write_text(json.dumps(changes), encoding="utf-8")
        return p

    def test_apply_basic(self, simple_5para_path, tmp_path, capsys):
        changes = [
            {
                "fragment_id": 1,
                "change_type": "modify",
                "new_text": "Modified first paragraph.",
                "justification": "Test change.",
            }
        ]
        changes_path = self._write_changes(tmp_path, changes)
        output_path = tmp_path / "output.docx"

        rc = main(["apply", str(simple_5para_path), str(changes_path), "-o", str(output_path)])
        assert rc == 0
        assert output_path.exists()
        assert output_path.stat().st_size > 0

    def test_apply_default_output_name(self, simple_5para_path, tmp_path, capsys):
        changes = [
            {
                "fragment_id": 2,
                "change_type": "delete",
                "justification": "Test delete.",
            }
        ]
        changes_path = self._write_changes(tmp_path, changes)

        rc = main(["apply", str(simple_5para_path), str(changes_path)])
        assert rc == 0
        expected_output = simple_5para_path.with_stem(simple_5para_path.stem + "_redlined")
        assert expected_output.exists()
        # Clean up
        expected_output.unlink()

    def test_apply_with_dict_wrapper(self, simple_5para_path, tmp_path, capsys):
        """Changes JSON wrapped in {"changes": [...]}."""
        changes = {
            "changes": [
                {
                    "fragment_id": 3,
                    "change_type": "append_after",
                    "new_text": "A new paragraph.",
                    "justification": "Test append.",
                }
            ]
        }
        changes_path = self._write_changes(tmp_path, changes)
        output_path = tmp_path / "output.docx"

        rc = main(["apply", str(simple_5para_path), str(changes_path), "-o", str(output_path)])
        assert rc == 0
        assert output_path.exists()

    def test_apply_missing_input(self, tmp_path, capsys):
        changes_path = self._write_changes(tmp_path, [])
        rc = main(["apply", "/nonexistent/input.docx", str(changes_path)])
        assert rc == 1

    def test_apply_invalid_json(self, simple_5para_path, tmp_path, capsys):
        bad_json = tmp_path / "bad.json"
        bad_json.write_text("not valid json {{{", encoding="utf-8")

        rc = main(["apply", str(simple_5para_path), str(bad_json)])
        assert rc == 1
        err = capsys.readouterr().err
        assert "Error" in err

    def test_apply_invalid_fragment_id(self, simple_5para_path, tmp_path, capsys):
        changes = [
            {
                "fragment_id": 999,
                "change_type": "modify",
                "new_text": "Text.",
                "justification": "Test.",
            }
        ]
        changes_path = self._write_changes(tmp_path, changes)
        output_path = tmp_path / "output.docx"

        rc = main(["apply", str(simple_5para_path), str(changes_path), "-o", str(output_path)])
        assert rc == 1

    def test_apply_no_validate(self, simple_5para_path, tmp_path, capsys):
        changes = [
            {
                "fragment_id": 1,
                "change_type": "modify",
                "new_text": "Modified.",
                "justification": "Test.",
            }
        ]
        changes_path = self._write_changes(tmp_path, changes)
        output_path = tmp_path / "output.docx"

        rc = main(
            [
                "apply",
                str(simple_5para_path),
                str(changes_path),
                "-o",
                str(output_path),
                "--no-validate",
            ]
        )
        assert rc == 0


class TestCliValidate:
    def test_validate_clean_file(self, simple_5para_path, capsys):
        rc = main(["validate", str(simple_5para_path)])
        assert rc == 0
        out = capsys.readouterr().out
        assert "passed" in out.lower()

    def test_validate_redlined_file(self, simple_5para_path, tmp_path, capsys):
        # First create a redlined file
        changes = [
            {
                "fragment_id": 1,
                "change_type": "modify",
                "new_text": "Modified text.",
                "justification": "Test.",
            }
        ]
        changes_path = tmp_path / "changes.json"
        changes_path.write_text(json.dumps(changes), encoding="utf-8")
        output_path = tmp_path / "redlined.docx"
        main(["apply", str(simple_5para_path), str(changes_path), "-o", str(output_path)])

        # Now validate it
        rc = main(["validate", str(output_path)])
        assert rc == 0

    def test_validate_missing_file(self, capsys):
        rc = main(["validate", "/nonexistent/file.docx"])
        assert rc == 1
