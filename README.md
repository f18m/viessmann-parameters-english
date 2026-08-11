# viessmann-parameters-english

List of parameters for Viessmann heat pumps and such, translated in English.

This repository contains my own translation (from German) of lists of Viessmann parameters that I found online.
The reason for this repository to exist is to enable non-German-speakers to 
understand the register mappings of their Viessmann devices.

## Sources of Viessmann parameters

* [openv](https://github.com/openv/openv/wiki): OpenV project -- the wiki is a good source of information but it's entirely written in German
* [philippo66 processed Excel](https://github.com/philippoo66/optolink-splitter/wiki/310-Parameter-Addresses): This is possibly the best, organized source of Viessmann parameter addresses but again all parameters are written in German. 
* [philippo66 ViessData](https://github.com/philippoo66/ViessData21): This is another great website containing dumps of data coming from the VitoSoft software package.


## Translated data

Translation is still work in progress:

* [Parameter Addresses](Viessmann_Parameter_Adresses_EN_translated.xlsx): this is the direct translation of  [philippo66 processed Excel](https://github.com/philippoo66/optolink-splitter/wiki/310-Parameter-Addresses)



## Viessmann Parameter Translator

My approach to translation consists in a small Python script to automate the translation of the thousands of parameters documented in the "Sources" repositories.

The translator uses a mix of static dictionaries and [DeepL API interface](https://www.deepl.com/it/translator) to achieve the semi-automated translation. 

### Files

- viessmann_translate.py  : translator
- dictionary.json          : technical German -> English overrides
- translation_review.csv   : unique tokens to review/correct
- README.txt               : this guide

### Recommended workflow

1. Open translation_review.csv in Excel.
2. Fill the EnglishTranslation column for technical terms where you want
   a specific Viessmann-style translation.
3. Put additional permanent overrides in dictionary.json.
4. Install dependencies:

    ```sh
    pip install openpyxl google-cloud-translate
    ```

5. Configure Google Cloud credentials:
     set DEEPL_API_KEY=.... 
   (PowerShell)
     $env:DEEPL_API_KEY="...."

6. Run:

    ```sh
    python viessmann_translate.py --input Viessmann_Parameter_Adresses_EN.xlsx ^
       --output Viessmann_Parameter_Adresses_EN_translated.xlsx ^
       --dictionary dictionary.json --review translation_review.csv ^
       --provider deepl
    ```

For a dry/local test without Google:

    ```sh
    python viessmann_translate.py --provider none
    ```

The program preserves hexadecimal addresses such as ~0xB490 and obvious
technical identifiers such as WPR3, HK2, KK2, etc.
