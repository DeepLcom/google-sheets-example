# DeepL API - Google Sheets Example (Sidebar UI)

This branch adds a **sidebar UI** to the Google Sheets DeepL integration. Select
cells in your sheet, open the DeepL sidebar, choose your translation options, and
click **Translate** — results are written back as plain text values, so they are
never re-translated when the sheet is reopened.

**Disclaimer**: the DeepL support team does *not* provide support for this
example project. If you have questions or feedback, please
[open a GitHub issue][issues].

## Requirements

- A **DeepL API authentication key**. [Create a free account here][pro-account]
  — the Free tier allows up to 500,000 characters/month.
- A **Google account** to use Google Sheets.

## Setup

### Option A — Copy the template sheet (recommended)

DeepL maintains a ready-to-use template Google Sheet with the script already
embedded. No coding required.

1. Open the [DeepL for Google Sheets template][template-sheet].
2. Click **File → Make a copy**. Give it a name and save it to your Drive.
3. Open your copy and click **DeepL → Open sidebar** in the toolbar.
4. Enter your DeepL API key when prompted. The sidebar verifies the key and
   takes you straight to the translation UI.

### Option B — Manual installation

Use this if you want to add DeepL to an existing sheet, or prefer to install
from source.

1. In your Google Sheet, go to **Extensions → Apps Script**.
2. Click **+** next to "Files", choose **Script**, name it `DeepL` and paste in the contents of
   [DeepL.gs][deepl-gs-raw].
3. Add another file, choose **HTML**, name it `DeepLSidebar`, and paste in the contents of
   [DeepLSidebar.html][deepl-sidebar-html-raw]. **Save the changes**.
4. Close the Apps Script tab and reload your sheet. A **DeepL** menu will
   appear in the toolbar.
5. Click **DeepL → Open sidebar** and enter your API key when prompted.

#### Updating

To check for updates, see the [CHANGELOG][changelog]. To update, open
**Extensions → Apps Script** and replace the contents of `DeepL.gs` and
`DeepLSidebar.html` with the latest versions (links above). Save and reload
your sheet. Your API key and saved options are stored in Script Properties and
are unaffected by updates.

## Usage

1. **Select** one or more cells containing the text you want to translate.
2. Click **DeepL → Open sidebar**.
3. Set your options:

   **Options** (always visible)

   | Option | Description |
   |---|---|
   | Source language | Language of the input text. Leave as Auto-detect if unsure. Choose *Other* to enter any BCP-47 code. |
   | Target language | Language to translate into. Choose *Other* to enter any BCP-47 code. |
   | Formality | Default / Formal / Informal. A note appears when the target language does not support this setting. |
   | Context | Optional hint to disambiguate the text (e.g. *"product listing for a luxury watch"*). Not translated. |

   **Customization** (collapsible)

   | Option | Description |
   |---|---|
   | Glossary ID | ID of a DeepL glossary to apply. |
   | Style rule ID | ID of a DeepL style rule list to apply. |
   | Custom instructions | Up to 10 plain-language instructions, one per line (e.g. *"Use gender-neutral language"*). |

   **Advanced** (collapsible)

   | Option | Description |
   |---|---|
   | Model | Default, Quality optimized, Prefer quality, or Speed optimized. |
   | Extra API options | Additional DeepL API parameters as `key=value` lines or a JSON object. |

4. Click **Translate**. Results are written back as plain text and will not be
   re-translated when the sheet is reopened.

All settings are saved and restored automatically between sessions.

The sidebar shows a **usage bar** for the current billing period, updated after
each translation. Your API key can be replaced or cleared at any time from the
**Settings** section at the bottom of the sidebar.

The installed version is shown in the sidebar footer and under **DeepL → About**.

## Formula functions

The script still includes the former `DeepLTranslate()` and `DeepLUsage()` spreadsheet
formula functions, for backward compatibility. These are **disabled by default** — read
[FORMULA_FUNCTIONS.md][formula-functions] for usage details, cost implications,
and instructions to enable them.

## Contributing

We welcome feedback and contributions. Please [open an issue][issues] or
[submit a pull request][pull-requests].

[api-languages]: https://www.deepl.com/docs-api/translating-text?utm_source=github&utm_content=google-sheets-plugin-readme&utm_medium=readme

[template-sheet]: https://docs.google.com/spreadsheets/d/1VVMDPYV7oL7ZM51RFUBDmmXnXRjYcMeiPx4gw5zAOgc/edit?usp=sharing

[deepl-gs-raw]: https://raw.githubusercontent.com/DeepLcom/google-sheets-example/main/DeepL.gs

[deepl-sidebar-html-raw]: https://raw.githubusercontent.com/DeepLcom/google-sheets-example/main/DeepLSidebar.html

[formula-functions]: FORMULA_FUNCTIONS.md

[changelog]: https://github.com/DeepLcom/google-sheets-example/blob/main/CHANGELOG.md

[issues]: https://github.com/DeepLcom/google-sheets-example/issues

[pull-requests]: https://github.com/DeepLcom/google-sheets-example/issues

[pro-account]: https://www.deepl.com/pro?utm_source=github&utm_content=google-sheets-plugin-readme&utm_medium=readme#developer

[cost-control]: https://support.deepl.com/hc/en-us/articles/360020685580-Cost-control
