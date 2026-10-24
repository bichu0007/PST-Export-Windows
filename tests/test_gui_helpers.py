from pathlib import Path

from gui import readable_name, unique_output_path


def test_readable_name_uses_relative_folder_path(tmp_path: Path):
    root = tmp_path / "mailboxes"
    source = root / "Department A" / "archive.pst"
    source.parent.mkdir(parents=True)
    source.touch()

    assert readable_name(source, root) == "Department_A__archive"


def test_readable_name_for_single_pst(tmp_path: Path):
    source = tmp_path / "archive.pst"
    source.touch()

    assert readable_name(source, tmp_path) == "archive"


def test_unique_output_path_does_not_overwrite(tmp_path: Path):
    first = unique_output_path(tmp_path, "archive")
    first.touch()
    second = unique_output_path(tmp_path, "archive")

    assert first != second
    assert second.parent == tmp_path
    assert second.suffix == ".csv"
