import pytest


def test_render() -> None:
    """
    Simple test to see if rendering succeeds
    :return:
    """

    os.environ["FOLDER_PATH"] = "./fixtures/"
    os.environ["OUTPUT_FILE"] = "./fixtures/render.docx"
    os.environ["AUTHOR_NAME"] = "My Name"
    os.environ["TITLE"] = "My Book"