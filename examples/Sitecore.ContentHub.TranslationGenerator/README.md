# Content Hub Portal Language Translations Generator

**Content Hub Portal Language Translations Generator** is a standalone console application that automates the translation of Content Hub language export files. Using [Microsoft Translator](https://learn.microsoft.com/en-us/azure/ai-services/translator/), it generates translated keys in your chosen target language, ready for re-import into Content Hub.

> ⚠️ **Note**
> The provided code is intended as a guideline and must be tailored to suit your specific implementation requirements. Please ensure thorough end-to-end testing is conducted to validate its functionality and performance in your environment.

## Prerequisites

* A [Content Hub](https://www.sitecore.com/products/content-hub) instance with exported language files ([guide](https://doc.sitecore.com/ch/en/users/content-hub/export-translations.html))
* An [Azure Translator resource](https://learn.microsoft.com/en-us/azure/ai-services/translator/how-to/create-translator-resource)

### Configuration

Set the following values in `appsettings.local.json`:

| Key                  | Description                                 |
| -------------------- | ------------------------------------------- |
| `Translation:ApiKey` | API key from your Azure Translator resource |
| `Translation:Region` | Region of your Azure Translator resource    |

## Usage

1. Build and run the console application.
2. When prompted, enter your **target language code** (e.g. `de-DE` for German).
3. Provide the path to your **Content Hub export file**.
4. Choose an output filename or press `Enter` to use the default
5. Repeat steps 2–4 for additional target languages as required.

## Importing Translations

Once generated, the translated file can be re-imported into Content Hub via the [standard import process](https://doc.sitecore.com/ch/en/users/content-hub/import-translations.html).