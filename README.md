# Excel to OJS3 XML conversion tool

## Licence
The conversion tool is distributed under the GNU GPL v2.

The conversion tool uses the PHPExcel library. PHPExcel is licensed under [LGPL (GNU LESSER GENERAL PUBLIC LICENSE)](https://github.com/PHPOffice/PHPExcel/blob/master/license.md)

## Changes in version 1.1.0.12 (Dec 2018)
- Support multilingual keywords

## Changes in version 1.1.0.11 (Nov 2018)
- Support remote galleys

## Changes in version 1.1.0.8 (Sep 2018)
- Support rich text in abstract fields

## Changes in version 1.1.0.7 (Sep 2018)
- Support for keywords and disciplines, authorEmail and authorMiddlename
- better support for articles in alternative locales, see below

## Todo
PHPExcel is deprecated, [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) should be used instead.

## Usage 
The tool was created for "in-house use" at the Federation of Finnish Learned Societies (https://tsv.fi). *It is not pretty*. It has not been thoroughly tested, but has been used to import the archives of six journals during 2017. Feel free to use and develop further.

Before importing the created data to your production server, you should try to import the data to a test environment to ensure that the created XML files work as expected. 

1. Create an Excel file containing the article data. See the details below and the "example.xlsx" file. The metadata of each article is in one row. The order of the columns does not matter. 
2. Order the Excel file according to the publication date field and the article sequence field
3. Move the Excel file to the same folder with the conversion script. Move the full text files to folder below the conversion script.
4. Edit "convert.php" file and change the settings in the beginning to match your needs.
5. Run "php convert.php". The script will create one XML per year.

## Article
| Field | Description |  Required|
|----------|:--------:|:--------:|
| prefix |  "The", "A" |  |
| title |  Article title | x |
| subTitle |  Article subtitle |   |
| abstract|  Article abstract |   |
| seq |  Article sequence inside an issue, first article '1' | x  |
| pages| For example "23-45"  |  |
| language| Alternative article language "en", "fi", "sv", "de", "fr"  |  |
| keywords| Word 1; Word 2; Word3 |  |
| disciplines| History; Political science; Astronomy |  |

## Issue
| Field | Description |  Required|
|----------|:--------:|:--------:|
| issueDatepublished |  Issue publication date, yyyy-mm-dd | x |
| issueVolume |  Issue volume | x |
| issueNumber |  Issue number | x |
| issueYear |  Issue year | x |
| issueTitle |  Issue title |  |
| sectionTitle |  Section title, eg. "Articles" | x  |
| sectionAbbrev |  Section abbreviation, eg. "ART" | x  |                    
                    
## Multiplied fields
An article can have multiple authors or full text files. Every article has to have at least one author and one file.

If an article has for example three authors, the excel file should include columns for each author with the number behind the column name changing. The first name of the third author will be saved to a field called *authorFirstname3*.

### Authors
| Field | Description |  Required|
|----------|:--------:|:--------:|
| authorFirstname1|  First name | x |
| authorMiddlename1|  Middle name |  |
| authorLastname1|  Last name | x  |
| authorEmail1|  Email |  |
| authorAffiliation1|  Affiliation |   |
| country1|  "FI", "SE", "DK", "CA", "US" |   |
| orcid1|  Orcid ID, should include "https://". Note that adding Orcid ID's this way is not recommended by Orcid. |   |
| authorBio1|  Biography |   |

### Files
| Field | Description |  Required|
|----------|:--------:|:--------:|
| file1|  Name of the file, "article1.pdf" or url for remote galley| x |
| fileLabel1|  Usually "PDF"| x |
| fileGenre1|  Usually "Article Text"| x |
| fileLocale1|  "en", "fi" etc. | x |

## Importing multilingual data

The new version of the converter supports three different ways of handling locales:
- Alternative 1: If all your data is in one language, you can just give the defaultLocale value in the converter settings.
- Alternative 2: If some of your articles are for example in English and some in Finnish, you can add an additional column named "language" and give the article locale in that column. See the example xls-file. All the article medata will be saved using the locale given in the language field. For example *title* can contain both English and Finnish titles as long as the language column matches the language used in the field.
- Alternative 3: If your articles are all in one language, but you also have some metadata in other languages, for example an abstract in another language, you can give an additional abstract field in a column named locale:abstract (for example en:abstract)


fi - Finnish
en - English
sv - Swedish
fr - French
de - German

