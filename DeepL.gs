/* Please change the line below */
const auth_key = "b493b8ef-0176-215d-82fe-e28f182c9544:fx"; // Replace with your authentication key

/**
 * Translates from one language to another using the DeepL Translation API.
 *
 * Note that you need to set your DeepL auth key by calling DeepLAuthKey() before use.
 *
 * @param {"Hello"} input The text to translate.
 * @param {"en"} source_lang Optional. The language code of the source language.
 *   Use "auto" to auto-detect the language.
 * @param {"es"} target_lang The language code of the target language.
 * @param {"def3a26b-3e84-..."} glossary_id Optional. The ID of a glossary to use
 *   for the translation.
 * @return Translated text.
 * @customfunction
 */
function DeepLTranslate(input, source_lang, target_lang, glossary_id) {
    if (!target_lang) target_lang = selectDefaultTargetLang_();
    let formData = {
        'target_lang': target_lang,
        'text': input
    };
    if (source_lang && source_lang !== 'auto') {
        formData['source_lang'] = source_lang;
    }
    if (glossary_id) {
        formData['glossary_id'] = glossary_id;
    }
    return httpRequest_('post', '/v2/translate', formData).translations[0].text;
}

/**
 * Retrieve information about your DeepL API usage during the current billing period.
 * @param {"count", "limit"} type Optional, retrieve the current used amount ("count")
 *   or the maximum allowed amount ("limit").
 * @return String explaining usage, or count or limit values as specified by type argument.
 * @customfunction
 */
function DeepLUsage(type) {
    const json = httpRequest_('get', '/v2/usage');
    const charCount = json.character_count;
    const charLimit = json.character_limit;
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
    const target_langs = [
        'bg', 'cs', 'da', 'de', 'el', 'en-gb', 'en-us', 'es', 'et', 'fi', 'fr', 'hu', 'id',
        'it', 'ja', 'lt', 'lv', 'nl', 'pl', 'pt-br', 'pt-pt', 'ro', 'ru', 'sk', 'sl', 'sv',
        'tr', 'zh'];
    const locale = Session.getActiveUserLocale().replace('_', '-').toLowerCase();
    if (target_langs.findIndex(locale) !== -1) return locale;
    const localePrefix = locale.substring(0, 2);
    if (target_langs.findIndex(localePrefix) !== -1) return localePrefix;
    if (localePrefix === 'en') return 'en-US';
    if (localePrefix === 'pt') return 'en-PT';
    return 'en';
}

/**
 * Helper function to execute HTTP requests.
 */
function httpRequest_(method, relative_url, formData = null) {
    const baseUrl = auth_key.endsWith(':fx')
        ? 'https://api-free.deepl.com'
        : 'https://api.deepl.com';
    const url = baseUrl + relative_url;
    const options = {
        method: method,
        muteHttpExceptions: true,
        headers: {
            'Authorization': 'DeepL-Auth-Key ' + auth_key,
        },
    };
    if (formData) options.payload = formData;
    const response = UrlFetchApp.fetch(url, options);
    let jsonContent = null;
    try {
        jsonContent = JSON.parse(response.getContentText());
    } catch (error) {
        throw new Error('Error occurred while parsing response: ' + response.getContentText());
    }
    const responseCode = response.getResponseCode();
    if (responseCode === 200) return jsonContent;
    const message = jsonContent.message || '';
    throw new Error('Error occurred while accessing the DeepL API: HTTP ' + responseCode + ' ' + message);
}
