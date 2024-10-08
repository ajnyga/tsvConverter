# Excel to OJS3 XML conversion tool

Version 1.6.0.0 supports the schema for OJS 3.3. (tested with OJS 3.3.0-17, Oct 2024)

The tool was originally created for "in-house use" at the Federation of Finnish Learned Societies (https://tsv.fi). The current version consitutes a major revision and includes new features. Feel free to use and develop further.

## Installation

This tool requires PHP 8.2 or greater.

Download and unzip the tsvConverter.

Make sure you can run php from command line.

Go to the tsvConverter folder and install or update dependencies via Composer (https://getcomposer.org/). The conversion tool uses https://github.com/PHPOffice/PhpSpreadsheet for reading sheets.

    composer install

## Usage 

Before importing the created data to your production server, **you should try to import the data to a test environment to ensure that the created XML files work as expected**.

Usage:

	php convert.php [options] -x <xlsx filename> -f <files folder name>

	Options:
	[--defaultLocale | -l] <4-digit locale code>
	[--onlyValidate | -v]

Convert:

	php convert.php -x sheetFilename -f filesFolderName

Only validate by adding -v:

	php convert.php -x sheetFilename -f filesFolderName -v


### Step by step instructions
1. Create an Excel file containing the article data. See the details below and the "exampleMinimal.xlsx" and "exampleAdvanced.xlsx" files. The metadata of each article is in one row. The order of the columns does not matter. 
2. Move the Excel file to the same folder as the conversion script. Move the full text files to a folder, for example "exampleFiles", below the conversion script.
3. Verify default values set in the file `config.ini`. In particular defaultLoacle (if not set via cli) and defaultUserGroupRef (see below).
4. Run `php convert.php -x exampleMinimal.xlsx -f exampleFiles`

Note that simple fields like, e.g. <description> can be added as columns to the excel sheet and will be converted to appropriate XML tags even if not listed in the tables below (see Advanced usage below).

The `defaultUserGroupRef` must be set in the file `config.ini` and needs to be compatible with the one used in your system (in the primary locale). Note that some journals (even with English as their primary language) may have a proprietary name for this group.

For larger imports it might be necessary to temporarily increase your OJS servers “post_max_size” and “upload_max_filesize” in your php.ini.

## Article
| Field | Description |  Required| Multilingual Support|
|----------|:--------:|:--------:|:--------:|
| prefix | "The", "A" |  | x |   |
| title | Article title | x | x |
| subtitle | Article subtitle |   | x |
| abstract | Article abstract |   | x |
| articleSeq |  Article sequence inside an issue, first article '1' | x  |   |
| pages | For example "23-45"  |  |   |
| language | Article language "en", "fi", "sv", "de", "fr"  | x |   |
| keywords | Word 1; Word 2; Word3 |  | x |
| disciplines | History; Political science; Astronomy |  | x |
| subjects | Subject1; Subject2; ... |  | x |
| articleCopyrightYear | 2005 |  |   |
| articleCopyrightHolder | "John Doe" |  |   |
| articleLicenseUrl | http://creativecommons.org/licenses/by/4.0 |  |   |
| articlePrimaryContactId  | Id of primary author (default = 1) |  |  |
| doi | "10.1234/art.182" |  |   |

## Issues & Sections
| Field | Description |  Required| Multilingual Support|
|----------|:--------:|:--------:|:--------:|
| issueDatePublished |  Issue publication date, yyyy-mm-dd. Note! has to be unique for each individual issue. | x |   |
| issueVolume |  Issue volume |  |   |
| issueNumber |  Issue number |  |   |
| issueYear |  Issue year | x |   |
| issueTitle |  Issue title |  | x |
| sectionTitle |  Section title, eg. "Articles" | x  | x |
| sectionAbbrev |  Section abbreviation, eg. "ART" | x  |   |
| sectionSeq |  Section sequence inside an issue, first section '1' |   |   |

## Multiplied fields
An article can have multiple authors or full text files. Every article has to have at least one author and one file.

If an article has for example three authors, the excel file should include columns for each author with the number behind the column name changing. The first name of the third author will be saved to a field called *authorFirstname3*.

### Authors
| Field | Description |  Required| Multilingual Support|
|----------|:--------:|:--------:|:--------:|
| authorGivenname1|  Given name | x | x |
| authorMiddlename1|  Middle name |  |   |
| authorFamilyname1|  Family name |   | x |
| authorEmail1|  Email |  |   |
| authorAffiliation1|  Affiliation |   | x |
| authorCountry1|  "FI", "SE", "DK", "CA", "US" |   |   |
| authorOrcid1|  Orcid ID, should include "https://". Note that adding Orcid ID's this way is not recommended by Orcid. |   |   |
| authorBiography1|  Biography |   | x |

### Files & Galleys
| Field | Description |  Required| Multilingual Support|
|----------|:--------:|:--------:|:--------:|
| fileName1|  Name of the file, "article1.pdf" or url for remote galley| x |   |
| fileGenre1|  Usually "Article Text"| x |   |
| galleyLabel1|  Usually "PDF"| x | x |
| galleyLocale1|  "en", "fi" etc. | x |   |

## Importing multilingual data

The converter supports three different ways of handling locales:
- Alternative 1: If all your data is in one language, you can just give the defaultLocale value in the converter settings.
- Alternative 2: If some of your articles are for example in English and some in Finnish, you can add an additional column named "language" and give the article locale in that column. See the example xls-file. All the article medata will be saved using the locale given in the language field. For example *title* can contain both English and Finnish titles as long as the language column matches the language used in the field.
- Alternative 3: If your articles are all in one language, but you also have some metadata in other languages, for example an abstract in another language, you can give an additional abstract field in a column named locale:abstract (for example en:abstract)


fi - Finnish
en - English
sv - Swedish
fr - French
de - German
ru - Russian
no - Norwegian
da - Danish
es - Spanish

## Validating XML

Install `libxml2-util` if xmllint is not already available, e.g.: `apt-get install libxml2-utils`.

Run `xmllint --noout --schema native.xsd <xml file>` to vaildate against OJS 3.3 native xsd.

## Additional xsd and xml files

The files `native.xsd`, `pkp-native.xsd` and `importexport.xsd` were taken from the OJS 3.3 repo for validation purposes.
The file `OJS_3.3_Native_Sample.xml` was automatically generated from theses xsd files by means of the Oxygen XML Editor. It provides a template from which XML tag order and required locale attributes are deduced.

## Advanced usage

The algorythm will convert any column name that resolves to a valid PKP native XML tag even if not specifically handled in the code. I.e., fields like creator, description, publisher, source, sponsor can be included in the excel sheet (even with locale codes). The default locale code is not used in this case. Column names without locale qualifier will have no locale attribute.

E.g. to add an issue description simply add a column `issueDescription`, or in Finish language create a column `fi:issueDescription`. A column `fileDescription2` will be interpreted as the tag `<description>` of the second `<submission_file>` tag.

## Licence
The conversion tool is distributed under the GNU GPL v3.

## Changes in version 1.6.0.0 (Oct 2024)
- revised command line parsing
- revised field/column naming
- full issue data used for issue identification
- rewrite to use PHP DOM model
- various XSDs for OJS 3.3 native XML added to improve XML handling and enable validation (by external tools like e.g. xmllint)
- config file added to handle default values

## Changes in version 1.5.0.0 (Aug 2023)
- new fields added

## Changes in version 1.4.0.0 (Aug 2023)
- Support OJS 3.3

## Changes in version 1.3.1.0 (Mar 2021)
- Support OJS 3.2

## Changes in version 1.2.0.0 (Mar 2021)
- Use PhpSpreadsheet (https://github.com/PHPOffice/PhpSpreadsheet) and Composer
- Use GPL v3

## Changes in version 1.1.0.12 (Dec 2018)
- Support multilingual keywords

## Changes in version 1.1.0.11 (Nov 2018)
- Support remote galleys

## Changes in version 1.1.0.8 (Sep 2018)
- Support rich text in abstract fields

## Changes in version 1.1.0.7 (Sep 2018)
- Support for keywords and disciplines, authorEmail and authorMiddlename
- better support for articles in alternative locales

