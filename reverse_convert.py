from __future__ import annotations

from app import TAB_FILE_TO_MARKDOWN, run


def main() -> None:
    run(initial_tab=TAB_FILE_TO_MARKDOWN)


if __name__ == "__main__":
    main()
