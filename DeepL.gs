/*
MIT License

Copyright 2022 DeepL SE (https://www.deepl.com)

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
 */

/* Set to true to enable the DeepLTranslate() and DeepLUsage() formula functions.
   Read FORMULA_FUNCTIONS.md for usage details and cost implications before enabling. */
const enableFormulaFunctions = false;

/* Change the line below to disable all translations. */
const disableTranslations = false; // Set to true to stop translations.

/* Change the line below to activate auto-detection of re-translations. */
const activateAutoDetect = false; // Set to true to enable auto-detection of re-translation.

/* You shouldn't need to modify the lines below here */

/* Version of this script from https://github.com/DeepLcom/google-sheet-example, included in logs. */
const scriptVersion = "0.4.0";

/**
 * Creates the DeepL menu when the spreadsheet is opened.
 */
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('DeepL')
        .addItem('Open sidebar', 'showSidebar')
        .addSeparator()
        .addItem('About (v' + scriptVersion + ')', 'showAbout_')
        .addToUi();
}

/**
 * Shows a dialog with the current script version and a link to the changelog.
 */
function showAbout_() {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
        'DeepL for Google Sheets',
        'Version ' + scriptVersion + '\n\n' +
        'To check for updates, visit:\n' +
        'https://github.com/DeepLcom/google-sheets-example/blob/main/CHANGELOG.md',
        ui.ButtonSet.OK
    );
}

/**
 * Returns the current script version. Called from the sidebar on load.
 * @return {string}
 */
function getScriptVersion() {
    return scriptVersion;
}

/**
 * Opens the DeepL translation sidebar.
 */
function showSidebar() {
    const html = HtmlService.createHtmlOutputFromFile('DeepLSidebar')
        .setTitle('DeepL Translate')
        .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Translates the currently selected cells and writes results back as static values.
 * Called from the sidebar via google.script.run.
 *
 * @param {string|null} sourceLang Source language code, or null for auto-detect.
 * @param {string} targetLang Target language code.
 * @param {{glossaryId, formality, context, styleId, customInstructions, modelType}} options
 * @return {{translated: number, skipped: number, failed: number, error: string, billedCharacters: number}}
 */
function translateSelectionFromSidebar(sourceLang, targetLang, options) {
    const range = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange();
    if (!range) throw new Error('No cells selected.');

    options = options || {};

    const flatCells = [];
    for (let r = 0; r < range.getNumRows(); r++) {
        for (let c = 0; c < range.getNumColumns(); c++) {
            flatCells.push(range.getCell(r + 1, c + 1));
        }
    }

    PropertiesService.getScriptProperties().setProperties({
        'DEEPL_LAST_SOURCE_LANG':         sourceLang                   || '',
        'DEEPL_LAST_TARGET_LANG':         targetLang,
        'DEEPL_LAST_FORMALITY':           options.formality            || '',
        'DEEPL_LAST_CONTEXT':             options.context              || '',
        'DEEPL_LAST_GLOSSARY_ID':         options.glossaryId           || '',
        'DEEPL_LAST_STYLE_ID':            options.styleId              || '',
        'DEEPL_LAST_CUSTOM_INSTRUCTIONS': options.customInstructions
                                            ? options.customInstructions.join('\n') : '',
        'DEEPL_LAST_MODEL_TYPE':          options.modelType            || '',
        'DEEPL_LAST_EXTRA_OPTIONS':       options.extraOptions         || '',
    });

    const cellsToTranslate = flatCells.filter(cell => {
        const text = cell.getDisplayValue();
        return text && text.trim() !== '';
    });
    const skipped = flatCells.length - cellsToTranslate.length;

    if (cellsToTranslate.length === 0) {
        return { translated: 0, skipped, failed: 0, error: '', billedCharacters: 0 };
    }

    try {
        const texts = cellsToTranslate.map(cell => cell.getDisplayValue());
        const results = callDeeplTranslateApi_(texts, sourceLang, targetLang, options);
        const billedCharacters = results.reduce((sum, r) => sum + r.billedCharacters, 0);
        for (let i = 0; i < cellsToTranslate.length; i++) {
            cellsToTranslate[i].setValue(results[i].text);
        }
        return { translated: cellsToTranslate.length, skipped, failed: 0, error: '', billedCharacters };
    } catch (e) {
        const lastError = e.message || String(e);
        Logger.log(`DeepLcom/google-sheets-example/${scriptVersion}: translateSelectionFromSidebar error: ${lastError}`);
        return { translated: 0, skipped, failed: cellsToTranslate.length, error: lastError, billedCharacters: 0 };
    }
}

/**
 * Returns API usage as a structured object for display in the sidebar.
 * Called from the sidebar via google.script.run.
 * @return {{charCount: number, charLimit: number}}
 */
function getUsageForSidebar() {
    const response = httpRequestWithRetries_('get', '/v2/usage');
    checkResponse_(response);
    const obj = JSON.parse(response.getContentText());
    if (obj.character_count === undefined || obj.character_limit === undefined)
        throw new Error('Character usage not found in API response.');
    return { charCount: obj.character_count, charLimit: obj.character_limit };
}

/**
 * Returns all saved sidebar options, or null if none have been saved yet.
 * Called from the sidebar on load.
 * @return {{sourceLang, targetLang, formality, context, glossaryId, styleId, customInstructions, modelType}|null}
 */
function getSavedOptions() {
    const props = PropertiesService.getScriptProperties();
    const targetLang = props.getProperty('DEEPL_LAST_TARGET_LANG');
    if (!targetLang) return null;
    return {
        sourceLang:         props.getProperty('DEEPL_LAST_SOURCE_LANG')          || '',
        targetLang,
        formality:          props.getProperty('DEEPL_LAST_FORMALITY')            || '',
        context:            props.getProperty('DEEPL_LAST_CONTEXT')              || '',
        glossaryId:         props.getProperty('DEEPL_LAST_GLOSSARY_ID')          || '',
        styleId:            props.getProperty('DEEPL_LAST_STYLE_ID')             || '',
        customInstructions: props.getProperty('DEEPL_LAST_CUSTOM_INSTRUCTIONS')  || '',
        modelType:          props.getProperty('DEEPL_LAST_MODEL_TYPE')           || '',
        extraOptions:       props.getProperty('DEEPL_LAST_EXTRA_OPTIONS')        || '',
    };
}

/**
 * Returns the last 4 characters of the saved API key, for display in the sidebar.
 * @return {string|null}
 */
function getApiKeySuffix() {
    const key = getApiKey_();
    return key ? key.slice(-4) : null;
}

/**
 * Returns whether an API key is currently saved in Script Properties.
 * Called from the sidebar on load.
 * @return {boolean}
 */
function hasApiKey() {
    return !!PropertiesService.getScriptProperties().getProperty('DEEPL_API_KEY');
}

/**
 * Saves the given API key to Script Properties and reloads the cached constant.
 * Called from the sidebar via google.script.run.
 * @param {string} key
 */
function saveApiKey(key) {
    if (!key || !key.trim()) throw new Error('API key must not be empty.');
    PropertiesService.getScriptProperties().setProperty('DEEPL_API_KEY', key.trim());
}

/**
 * Deletes the saved API key from Script Properties.
 * Called from the sidebar via google.script.run.
 */
function clearApiKey() {
    PropertiesService.getScriptProperties().deleteProperty('DEEPL_API_KEY');
}

/**
 * Calls the DeepL translate API for a single text string.
 * Shared by the formula function and the sidebar/menu actions.
 *
 * @param {string} text The text to translate.
 * @param {string|null} sourceLang Source language code, or null/falsy for auto-detect.
 * @param {string} targetLang Target language code.
 * @param {string|null} glossaryId Glossary ID, or null to omit.
 * @return {string} Translated text.
 */
function callDeeplTranslateApi_(texts, sourceLang, targetLang, options) {
    options = options || {};
    const body = { text: texts, target_lang: targetLang, show_billed_characters: true };
    if (sourceLang)                                 body.source_lang          = sourceLang;
    if (options.glossaryId)                         body.glossary_id          = options.glossaryId;
    if (options.formality)                          body.formality            = options.formality;
    if (options.context)                            body.context              = options.context;
    if (options.styleId)                            body.style_id             = options.styleId;
    if (options.customInstructions && options.customInstructions.length)
                                                    body.custom_instructions  = options.customInstructions;
    if (options.modelType)                          body.model_type           = options.modelType;
    if (options.extraOptions) {
        const raw = options.extraOptions.trim();
        if (raw.startsWith('{')) {
            let parsed;
            try { parsed = JSON.parse(raw); } catch (e) {
                throw new Error('Extra options: invalid JSON — ' + e.message);
            }
            Object.assign(body, parsed);
        } else {
            for (const line of raw.split('\n')) {
                const idx = line.indexOf('=');
                if (idx > 0) {
                    const key = line.slice(0, idx).trim();
                    const val = line.slice(idx + 1).trim();
                    if (key) body[key] = val;
                }
            }
        }
    }
    const totalChars = texts.reduce((sum, t) => sum + t.length, 0);
    const response = httpRequestWithRetries_('post', '/v2/translate', body, totalChars, true);
    checkResponse_(response);
    return JSON.parse(response.getContentText()).translations
        .map(t => ({ text: t.text, billedCharacters: t.billed_characters || 0 }));
}

function getApiKey_() {
    return PropertiesService.getScriptProperties().getProperty('DEEPL_API_KEY');
}

/**
 * Translates from one language to another using the DeepL Translation API.
 *
 * @param {"Hello"} input The text to translate.
 * @param {"en"} sourceLang Optional. The language code of the source language.
 *   Use "auto" to auto-detect the language.
 * @param {"es"} targetLang Optional. The language code of the target language.
 *   If unspecified, defaults to your system language.
 * @param {"def3a26b-3e84-..."} glossaryId Optional. The ID of a glossary to use
 *   for the translation.
 * @param {cell range} options Optional. Range of additional options to send with API translation
 *   request. May also be specified inline e.g. '{"tag_handling", "xml"; "ignore_tags", "ignore"}'
 * @return Translated text.
 * @customfunction
 */
function DeepLTranslate(input,
                        sourceLang,
                        targetLang,
                        glossaryId,
                        options
) {
    if (!enableFormulaFunctions) {
        throw new Error('Formula functions are disabled. Set enableFormulaFunctions = true in DeepL.gs to enable them. See FORMULA_FUNCTIONS.md for details and cost implications.');
    }
    if (input === undefined) {
        throw new Error("input field is undefined, please specify the text to translate.");
    } else if (typeof input === "number") {
        input = input.toString();
    } else if (typeof input !== "string") {
        throw new Error("input text must be a string.");
    }
    // Check the current cell to detect recalculations due to reopening the sheet
    const cell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getCurrentCell();

    if (disableTranslations) {
        Logger.log(`DeepLcom/google-sheets-example/${scriptVersion}: disableTranslations is active, skipping DeepL translation request`);
        return cell.getDisplayValue();
    }

    if (activateAutoDetect &&
            cell.getDisplayValue() !== "" &&
            cell.getDisplayValue() !== "Loading...") {
        Logger.log(`DeepLcom/google-sheets-example/${scriptVersion}: Detected cell-recalculation, skipping DeepL translation request`);
        return cell.getDisplayValue();
    }

    if (!targetLang) targetLang = selectDefaultTargetLang_();
    let formData = {
        'target_lang': targetLang,
        'text': input
    };
    if (sourceLang && sourceLang !== 'auto') {
        formData['source_lang'] = sourceLang;
    }
    if (glossaryId) {
        formData['glossary_id'] = glossaryId;
    }
    if (options) {
        if (!Array.isArray(options) ||
            !Object.values(options).every(function(value) {
                return Array.isArray(value) && value.length === 2;
            })) {
            throw new Error("options must be a range with two columns, or have the form '{\"opt1\", \"val1\"; \"opt2\", \"val2\"}'");
        }

        for (let i = 0; i < options.length; i++) {
            const items = options[i];
            const key = items[0];
            formData[key] = items[1];
        }
    }

    const response = httpRequestWithRetries_('post', '/v2/translate', formData, input.length);
    checkResponse_(response);
    const responseObject = JSON.parse(response.getContentText());
    return responseObject.translations[0].text;
}

/**
 * Retrieve information about your DeepL API usage during the current billing period.
 * @param {"count", "limit"} type Optional, retrieve the current used amount ("count")
 *   or the maximum allowed amount ("limit").
 * @return String explaining usage, or count or limit values as specified by type argument.
 * @customfunction
 */
function DeepLUsage(type) {
    if (!enableFormulaFunctions) {
        throw new Error('Formula functions are disabled. Set enableFormulaFunctions = true in DeepL.gs to enable them. See FORMULA_FUNCTIONS.md for details and cost implications.');
    }
    const response = httpRequestWithRetries_('get', '/v2/usage');
    checkResponse_(response);
    const responseObject = JSON.parse(response.getContentText());
    const charCount = responseObject.character_count;
    const charLimit = responseObject.character_limit;
    if (charCount === undefined || charLimit === undefined)
        throw new Error('Character usage not found.');
    if (type) {
        if (type === 'count') return charCount;
        if (type === 'limit') return charLimit;
        throw new Error('Unrecognized type argument.');
    }
    return `${charCount} of ${charLimit} characters used.`;
}

/////////////////////////////////////////////////////////////////////////////////////////
// General helper functions
/////////////////////////////////////////////////////////////////////////////////////////

/**
 * Determines the default target language using the system language.
 * @return A DeepL-supported target language.
 * @throws Error If the system language could not be converted to a supported target language.
 */
function selectDefaultTargetLang_() {
    const targetLangs = [
        'bg', 'cs', 'da', 'de', 'el', 'en-gb', 'en-us', 'es', 'et', 'fi', 'fr', 'hu', 'id',
        'it', 'ja', 'lt', 'lv', 'nl', 'pl', 'pt-br', 'pt-pt', 'ro', 'ru', 'sk', 'sl', 'sv',
        'tr', 'zh'];
    const locale = Session.getActiveUserLocale().replace('_', '-').toLowerCase();
    if (targetLangs.indexOf(locale) !== -1) return locale;
    const localePrefix = locale.substring(0, 2);
    if (targetLangs.indexOf(localePrefix) !== -1) return localePrefix;
    if (localePrefix === 'en') return 'en-US';
    if (localePrefix === 'pt') return 'en-PT';
    return 'en';
}

/**
 * Helper function to check response code is OK and if not, throw useful exceptions.
 */
function checkResponse_(response) {
    const responseCode = response.getResponseCode();
    if (200 <= responseCode && responseCode < 400) return;

    const content = response.getContentText();

    let message = '';
    try {
        const jsonObj = JSON.parse(content);
        if (jsonObj.message !== undefined) {
            message += `, message: ${jsonObj.message}`;
        }
        if (jsonObj.detail !== undefined) {
            message += `, detail: ${jsonObj.detail}`;
        }
    } catch (error) {
        // JSON parsing errors are ignored, and we fall back to the raw content
        message = ', ' + content;
    }

    switch (responseCode) {
        case 403:
            throw new Error(`Authorization failure, check DeepL API key, ${message}`);
        case 456:
            throw new Error(`Quota for this billing period has been exceeded${message}`);
        case 400:
            throw new Error(`Bad request${message}`);
        case 429:
            throw new Error(
                `Too many requests, DeepL servers are currently experiencing high load${message}`,
            );
        default: {
            throw new Error(
                `Unexpected status code: ${responseCode} ${message}, content: ${content}`,
            );
        }
    }
}

/**
 * Helper function to execute HTTP requests and retry failed requests.
 */
function httpRequestWithRetries_(method, relativeUrl, formData = null, charCount = 0, useJson = false) {
    const apiKey = getApiKey_();
    if (!apiKey) {
        throw new Error('DeepL API key not set. Use the DeepL sidebar to add your API key.');
    }
    const baseUrl = apiKey.endsWith(':fx')
        ? 'https://api-free.deepl.com'
        : 'https://api.deepl.com';
    const url = baseUrl + relativeUrl;
    const params = {
        method: method,
        muteHttpExceptions: true,
        headers: {
            'Authorization': 'DeepL-Auth-Key ' + apiKey,
        },
    };
    if (formData) {
        if (useJson) {
            params.contentType = 'application/json';
            params.payload = JSON.stringify(formData);
        } else {
            params.payload = formData;
        }
    }
    let response = null;
    for (let numRetries = 0; numRetries < 5; numRetries++) {
        const lastRequestTime = Date.now();
        try {
            Logger.log(`DeepLcom/google-sheets-example/${scriptVersion}: Sending HTTP request to ${url} with ${charCount} characters`);
            response = UrlFetchApp.fetch(url, params);
            const responseCode = response.getResponseCode();
            if (responseCode !== 429 && responseCode < 500) {
                return response;
            }
        } catch (e) {
            // It would be sensible to check whether the exception is retryable here, but there is
            // not so much documentation on Google Apps Script exceptions. In addition, UrlFetchApp
            // fetch timeouts are very long and not configurable.
            throw e;
        }
        Logger.log(`DeepLcom/google-sheets-example/${scriptVersion}: Retrying after ${numRetries} failed requests.`);
        sleepForBackoff(numRetries, lastRequestTime);
    }
    return response;
}

/**
 * Helper function to sleep after failed requests.
 */
function sleepForBackoff(numRetries, lastRequestTime) {
    const backoff = Math.min(1000 * (1.6 ** numRetries), 60000);
    const jitter = 1 + 0.23 * (2 * Math.random() - 1); // Random value in [0.77 1.23]
    const sleepTime = Date.now() - lastRequestTime + backoff * jitter;
    Utilities.sleep(sleepTime);
}
