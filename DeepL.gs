/* Please change the line below */
const authKey = "b493b8ef-0176-215d-82fe-e28f182c9544:fx"; // Replace with your authentication key

/**
 * Translates from one language to another using the DeepL Translation API.
 *
 * Note that you need to set your DeepL auth key by calling DeepLAuthKey() before use.
 *
 * @param {"Hello"} input The text to translate.
 * @param {"en"} sourceLang Optional. The language code of the source language.
 *   Use "auto" to auto-detect the language.
 * @param {"es"} targetLang The language code of the target language.
 * @param {"def3a26b-3e84-..."} glossaryId Optional. The ID of a glossary to use
 *   for the translation.
 * @return Translated text.
 * @customfunction
 */
function DeepLTranslate(input, sourceLang, targetLang, glossaryId) {
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
    return charCount + ' of ' + charLimit + ' characters used.';
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
    if (targetLangs.findIndex(locale) !== -1) return locale;
    const localePrefix = locale.substring(0, 2);
    if (targetLangs.findIndex(localePrefix) !== -1) return localePrefix;
    if (localePrefix === 'en') return 'en-US';
    if (localePrefix === 'pt') return 'en-PT';
    return 'en';
}

/**
 * Helper function to check response code and if not, throw useful exceptions.
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
            throw new Error(`Authorization failure, check authKey${message}`);
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
function httpRequestWithRetries_(method, relativeUrl, formData = null, charCount = 0) {
    const baseUrl = authKey.endsWith(':fx')
        ? 'https://api-free.deepl.com'
        : 'https://api.deepl.com';
    const url = baseUrl + relativeUrl;
    const params = {
        method: method,
        muteHttpExceptions: true,
        headers: {
            'Authorization': 'DeepL-Auth-Key ' + authKey,
        },
    };
    if (formData) params.payload = formData;
    let response = null;
    for (let numRetries = 0; numRetries < 5; numRetries++) {
        const lastRequestTime = Date.now();
        try {
            Logger.log(`Sending HTTP request to ${url} with ${charCount} characters`);
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
        Logger.log(`Retrying after ${numRetries} failed requests.`);
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
