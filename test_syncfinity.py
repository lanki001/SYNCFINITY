# test_syncfinity.py
import os
import shutil
import tempfile
import pytest


# ─── SETUP: Create temporary source and destination folders for each test ────

@pytest.fixture
def folders():
    """Creates a fresh temp source and destination folder for every test.
    Cleans them up automatically after each test finishes."""
    src = tempfile.mkdtemp()
    dst = tempfile.mkdtemp()
    yield src, dst
    shutil.rmtree(src)
    shutil.rmtree(dst)


# ─── THE COPY FUNCTION (mirrors what SyncFinity does internally) ─────────────

def sync_files(source_folder, destination_folder):
    for filename in os.listdir(source_folder):
        src_path = os.path.join(source_folder, filename)
        dst_path = os.path.join(destination_folder, filename)
        if os.path.isfile(src_path):
            shutil.copy2(src_path, dst_path)


# ─── TESTS ───────────────────────────────────────────────────────────────────

def test_copied_file_exists_in_destination(folders):
    src, dst = folders
    test_file = os.path.join(src, "report.txt")
    with open(test_file, "w") as f:
        f.write("syncing this file")

    sync_files(src, dst)

    assert os.path.exists(os.path.join(dst, "report.txt")), \
        "File was not copied to destination"


def test_copied_file_content_matches(folders):
    src, dst = folders
    test_file = os.path.join(src, "data.txt")
    with open(test_file, "w") as f:
        f.write("hello from syncfinity")

    sync_files(src, dst)

    with open(os.path.join(dst, "data.txt"), "r") as f:
        content = f.read()

    assert content == "hello from syncfinity", \
        "File content in destination does not match source"

# def test_multiple_files_all_copied(folders):
#     """All files in source must appear in destination"""
#     src, dst = folders

#     filenames = ["file1.txt", "file2.txt", "file3.txt"]
#     for name in filenames:
#         with open(os.path.join(src, name), "w") as f:
#             f.write(f"content of {name}")

#     sync_files(src, dst)

#     for name in filenames:
#         assert os.path.exists(os.path.join(dst, name)), \
#             f"{name} was not copied to destination"


# def test_empty_source_folder_does_not_crash(folders):
#     """Syncing an empty folder should complete without errors"""
#     src, dst = folders
#     sync_files(src, dst)  # should not raise any exception


# def test_original_file_still_exists_after_copy(folders):
#     """Source file must still exist after sync — it's a copy, not a move"""
#     src, dst = folders

#     test_file = os.path.join(src, "original.txt")
#     with open(test_file, "w") as f:
#         f.write("do not delete me")

#     sync_files(src, dst)

#     assert os.path.exists(test_file), \
#         "Source file was deleted after sync — it should still be there"
