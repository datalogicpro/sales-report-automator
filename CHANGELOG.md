# Changelog

All notable changes to this project are listed here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project uses [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added

- A web page: try the sample file or upload your own, see the summaries and charts, and download the Excel report.
- A check for the 7 required columns, with a message naming any that are missing. Column names still match with different capitals or extra spaces.
- Clear error messages for a file that is missing, unreadable, empty, or the wrong type.
- An `--output` option to choose where the report is saved.
- Logging for each step, in place of printed messages.
- Tests with 100% coverage, Ruff, mypy, pip-audit, and pre-commit.
- GitHub Actions for the checks and tests, and Dependabot.

### Changed

- Dependencies are managed with uv in `pyproject.toml` and `uv.lock`, in place of `requirements.txt`.

### Removed

- Support for old `.xls` Excel files, which needed a library the project didn't include. `.xlsx` files still work.
