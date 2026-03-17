# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).


## [Unreleased]

## [0.3.0] - 2026-03-17
### Added
* **Sidebar UI for translation**. Select cells, set options, and click
  **Translate** — results are written as static values and never
  re-translated when the sheet is reopened.
* Translation options: source/target language, formality
  (Default / Formal / Informal), context hint, and glossary ID. All
  options are saved and restored automatically between sessions.
* API key management in the sidebar: enter, validate, and clear the key
  without opening Apps Script settings.
* Usage bar showing character consumption for the current billing period,
  updated after each translation. Billed characters for the current
  translation are shown after translation is complete.

### Changed
* The `DeepLTranslate()` and `DeepLUsage()` functions are maintained for
  backwards compatibility, but disabled by default, see 
  `FORMULA_FUNCTIONS.md` for instructions.

## [0.2.0] - 2025-09-01
### Changed
* Renamed `freeze` variable to `disableTranslations`.
* Re-translation auto-detection is now off by default, and can be activated by
  `activateAutoDetect` variable.
* Moved `DEEPL_API_KEY` to script properties, for ease of use. Additionally,
  including the API key directly into the script is not very secure.


## [0.1.0] - 2022-07-06
Initial release.


[Unreleased]: https://github.com/DeepLcom/google-sheets-example/compare/v0.3.0...HEAD
[0.3.0]: https://github.com/DeepLcom/google-sheets-example/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/DeepLcom/google-sheets-example/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/DeepLcom/google-sheets-example/releases/tag/v0.1.0
