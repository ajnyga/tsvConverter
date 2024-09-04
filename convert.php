<?php

# Usage:
# php convert.php sheetFilename filesFolderName

// PHPExcel settings
// error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/Helsinki');
define('EOL', (PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ConvertExcel2PKPNativeXML {
	
	// cli parsing
	private $opts;
	private $posArgs;
	private $filesFolder;
	private $onlyValidate = false;
	private $fileName = 'articleData.xlsx';
	private $files = 'files';

	// defaults
	private $defaultUploader = 'admin';
	private $defaultAuthor;
	private $defaultUserGroupRef;
	private $defaultLocale = 'en_US';

	// table parsing
	private $issueKeys;
	private $sectionKeys;
	private $articleKeys;
	private $locales = array(
		'en' => 'en_US',
		'fi' => 'fi_FI',
		'sv' => 'sv_SE',
		'de' => 'de_DE',
		'ge' => 'de_DE',
		'ru' => 'ru_RU',
		'fr' => 'fr_FR',
		'no' => 'nb_NO',
		'da' => 'da_DK',
		'es' => 'es_ES',
	);

	// Constructor
	public function __construct($argv) {

		// pasre cli
		$rest_index = null;
		$shortOpts = "vl:x:f:";
		$longOpts = ['defaultLocale:', 'validate'];
		$this->opts = getopt($shortOpts, $longOpts, $rest_index);
		$this->posArgs = array_slice($argv, $rest_index);

		if (!$this->validateInput()) {

		}

		// set defaults

		// Default author name. If no author is given for an article, this name is used instead.
		$this->defaultAuthor['givenname'] = "Editorial Board";

		// Default user group (localized)
		$this->defaultUserGroupRef = array(
			'en_US' => 'Author',
			'de_DE' => 'Autor/in',
			'sv_SE' => 'F&#xF6;rfattare'
		);

		// load data
		echo date('H:i:s'), " Creating a new PHPExcel object", EOL;
		$objReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($this->fileName);
		$objReader->setReadDataOnly(false);
		$objPhpSpreadsheet = $objReader->load($this->fileName);
		$sheet = $objPhpSpreadsheet->setActiveSheetIndex(0);

		echo date('H:i:s'), " Creating an array", EOL;
		$articles = $this->createArray($sheet);
		$maxAuthors = $this->countMaxAuthors($sheet);
		$maxFiles = $this->countMaxFiles($sheet);

		/* 
		* Data validation   
		* -----------
		*/

		echo date('H:i:s'), " Validating data", EOL;

		$errors = $this->validateArticles($articles);
		if ($errors != "") {
			echo $errors, EOL;
			die();
		}

		# If only validation is selected, exit
		if ($this->onlyValidate == 1) {
			echo date('H:i:s'), " Validation complete ", EOL;
			die();
		}

		$this->process($articles);
	}

	function process($articles) {

		/* 
		* Prepare data for output
		* ----------------------------------------
		*/

		echo date('H:i:s'), " Preparing data for output", EOL;

		# Create issue and section identification

		$this->issueKeys = $this->getUniqueKeys($articles, 'issue');
		$this->sectionKeys = $this->getUniqueKeys($articles, 'section');
		$this->articleKeys = array_diff($this->getUniqueKeys($articles), $this->issueKeys, $this->sectionKeys);
		// $sections = [];
		$issueIdentifications = [];
		$issueData = [];
		//  reindex array (keys used as article id later)
		$articles = array_values($articles);
		foreach ($articles as $id => $article) {

			// identify issue and sections for each article by hash generated from issue and section fields
			$issueIdentification = array_intersect_key($article, array_flip($this->issueKeys));
			$article['articleIssueHash'] = hash("sha256", implode(", ", $issueIdentification));

			$sectionIdentification = array_intersect_key($article, array_flip($this->sectionKeys));
			$article['articleSectionHash'] = hash("sha256", implode(", ", $sectionIdentification));

			$issueIdentifications[$article['articleIssueHash']] = $issueIdentification;

			$issueData[$article['articleIssueHash']]['issue_identification'] = $issueIdentification;
			$issueData[$article['articleIssueHash']]['sections']['section'] = $sectionIdentification;

			$issueData[$article['articleIssueHash']]['articles'][$id] = array_intersect_key($article, array_flip($this->articleKeys));
			$issueData[$article['articleIssueHash']]['articles'][$id]['sectionAbbrev'] = $sectionIdentification['sectionAbbrev'];
		}

		/* 
		* Create XML  
		* --------------------
		*/

		echo date('H:i:s'), " Starting XML output", EOL;
		$currentIssueDatepublished = null;
		$currentYear = null;
		$submission_file_id = 1;
		$authorId = 1;
		$submissionId = 1;
		$file_id = 1;

		$dom = new DOMDocument('1.0', 'UTF-8');
		$dom->formatOutput = true;

		$issuesDOM = $dom->createElement('issues');
		$dom->appendChild($issuesDOM);

		// Create issue DOMs
		foreach ($issueData as $issueHash => $issueDataContent) {
			$issuesDOM = $this->processData($issuesDOM, $issueDataContent);
		}

		$dom->save('test.xml');
	}

	function validateInput() {
		// Check if the required parameter -x is set
		if (!isset($this->opts['x'])) {
			echo "Error: Required parameter -x <xlsx filename> is missing.\n";
			exit(1);
		}
		$this->fileName = $this->opts['x'];

		// Check if the required parameter -f is set
		$this->files = "files";
		if (!isset($this->opts['f'])) {
			echo "Error: Required parameter -f <files folder name> is missing.\n";
			exit(1);
		}
		$this->files = $this->opts['f'];

		// Check if the optional validate flag is set
		$this->onlyValidate = 0;
		if (isset($this->opts['v']) || isset($this->opts['validate'])) {
			echo "Validation only mode is enabled.\n";
			$this->onlyValidate = 1;
		}

		// Check if the defaultLocale optional parameter is set
		if (isset($this->opts['l'])) {
			echo "Default locale is set with value: " . $this->opts['l'] . "\n";
			$this->defaultLocale = $this->opts['l'];
		} elseif (isset($this->opts['defaultLocale'])) {
			echo "Default locale is set with value: " . $this->opts['defaultLocale'] . "\n";
			// The default locale. For alternative locales use language field. For additional locales use locale:fieldName.
			$this->defaultLocale = $this->opts['defaultLocale'];
		}

		/* 
		* Check that a file and a folder exists
		* ------------------------------------
		*/

		if (!file_exists($this->fileName)) {
			echo date('H:i:s') . " ERROR: Excel file does not exist" . EOL;
			die();
		}

		// Location of full text files
		$this->filesFolder = dirname(__FILE__) . "/" . $this->files . "/";

		if (!file_exists($this->filesFolder)) {
			echo date('H:i:s') . " ERROR: given folder does not exist" . EOL;
			die();
		}

		return true;
	}

	function processData($dom, $data) {
		foreach ($data as $tagName => $content) {
			// Create a new element with tag name and data
			switch ($tagName) {
				case 'sections':
					$sectionsDOM = $this->getOrCreateCollectionDOM($dom, 'sections');
					$sectionDOM = $dom->ownerDocument->createElement('section');
					$sectionsDOM->appendChild($sectionDOM);
	
					$sectionDOM = $this->processData($sectionDOM, $data['sections']['section']);
	
					break;
				case 'issue_identification':
					$issueDOM = $dom->ownerDocument->createElementNS('http://pkp.sfu.ca', 'issue');
					$issueDOM->setAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');
					$issueDOM->setAttribute('xsi:schemaLocation', 'http://pkp.sfu.ca native.xsd');
					$issueDOM->setAttribute('published', '1');
					$issueDOM->setAttribute('current', '0');
					$dom->appendChild($issueDOM);
					
					$issuesIdentificationDOM = $dom->ownerDocument->createElement($tagName);
					$issueDOM->appendChild($issuesIdentificationDOM );
					$issuesIdentificationDOM = $this->processData($issuesIdentificationDOM, $data[$tagName]);
					break;
				case 'sectionTitle':
				case 'sectionAbbrev':
					$element = $dom->ownerDocument->createElement(
						strtolower(str_replace('section', '', $tagName)),
						htmlspecialchars($data[$tagName])
					);
					$dom->setAttribute('ref', htmlspecialchars($data['sectionAbbrev']));
					$dom->setAttribute('seq', htmlspecialchars(isset($data['sectionSeq']) ? $data['sectionSeq'] : "0"));
					$dom->appendChild($element);
					break;
				case 'issueVolume':
				case 'issueNumber':
				case 'issueYear':
				case 'issueTitle':
					$element = $dom->ownerDocument->createElement(
						strtolower(str_replace('issue', '', $tagName)),
						htmlspecialchars($data[$tagName])
					);
					$dom->appendChild($element);
					break;
				case 'issueDatePublished':
					$element = $dom->ownerDocument->createElement('date_published', $content);
					$dom->appendChild($element);
					$element = $dom->ownerDocument->createElement('last_modified', $content);
					$dom->appendChild($element);
					break;
				case 'articles':
					$articlesDOM = $this->getOrCreateCollectionDOM($dom, 'articles', 'http://pkp.sfu.ca');
					$articlesDOM->setAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');
					$articlesDOM->setAttribute('xsi:schemaLocation', 'http://pkp.sfu.ca native.xsd');
	
					foreach ($data['articles'] as $id => $article) {
						$articleDOM = $dom->ownerDocument->createElement('article');
						$id = $dom->ownerDocument->createElement('id', $id);
						$id->setAttribute('type', 'internal');
						$id->setAttribute('advice', 'ignore');
						$articleDOM->appendChild($id);

						# Article
						echo date('H:i:s'), " Adding article: ", $article['title'], EOL;

						# Check if language has an alternative default locale
						# If it does, use the locale in all fields
						$articleLocale = $this->defaultLocale;
						if (!empty($article['language'])) {
							$articleLocale = $this->locales[trim($article['language'])];
						}

						$publicationDOM = $dom->ownerDocument->createElement('publication');
						$publicationDOM->setAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');
						$publicationDOM->setAttribute('xsi:schemaLocation', 'http://pkp.sfu.ca native.xsd');
						$publicationDOM->setAttribute('locale', $articleLocale);
						$publicationDOM->setAttribute('version', "1");
						$publicationDOM->setAttribute('status', "3");
						$publicationDOM->setAttribute('access_status', "0");
						// $publicationDOM->setAttribute('primary_contact_id', $authorId); TODO @RS
						$publicationDOM->setAttribute('url_path', "");
						$publicationDOM->setAttribute('date_published', $dom->getElementsByTagName('date_published')[0]->textContent);
						$publicationDOM->setAttribute('section_ref', $article['sectionAbbrev']);
						if (isset($article['articleSeq'])) {
							$publicationDOM->setAttribute('seq', $article['articleSeq']);
						}
						$articleDOM->appendChild($publicationDOM);

						$publicationDOM = $this->processData($publicationDOM, $article);

						// TODO @RS submission_file

						$articlesDOM->appendChild($articleDOM);
					}
	
					break;
				default:
					// here we handle all article columns
					if (in_array($tagName, $this->articleKeys)) {
						switch ($tagName) {
							case 'title':
								
								break;
						}
					}
					break;
			}
		}
		return $dom;
	}
	
	function getOrCreateCollectionDOM($dom, $tagname, $namespace = NULL) {
		$collectionDOM = $dom->ownerDocument->getElementById($tagname);
		if (!isset($collectionDOM) || $collectionDOM->length == 0) {
			if ($namespace) {
				$collectionDOM = $dom->ownerDocument->createElementNS($namespace, $tagname);
			} else {
				$collectionDOM = $dom->ownerDocument->createElement($tagname);
			}
			$dom->lastChild->appendChild($collectionDOM);
		}
		return $collectionDOM;
	}

	/* 
	* Helpers 
	* -----------
	*/


	# Function for searching alternative locales for a given field
	function searchLocalisations($key, $input, $intend, $tag = null, $flags = null)
	{
		global $locales;

		if ($tag == "") $tag = $key;

		$nodes = "";
		$pattern = "/:" . $key . "/";
		$values = array_intersect_key($input, array_filter(array_flip(preg_grep($pattern, array_keys($input), $flags ?? 0))));

		foreach ($values as $keyval => $value) {
			if ($value != "") {
				$shortLocale = explode(":", $keyval);
				if (strpos($value, "\n") !== false || strpos($value, "&") !== false || strpos($value, "<") !== false || strpos($value, ">") !== false) $value = "<![CDATA[" . nl2br($value) . "]]>";
				for ($i = 0; $i < $intend; $i++) $nodes .= "\t";
				$nodes .= "<" . $tag . " locale=\"" . $locales[$shortLocale[0]] . "\">" . $value . "</" . $tag . ">\r\n";
			}
		}

		return $nodes;
	}

	# Function for searching alternative locales for a given taxonomy field
	function searchTaxonomyLocalisations($key, $key_singular, $input, $intend, $flags = 0)
	{
		global $locales;

		$nodes = "";
		$intend_string = "";
		for ($i = 0; $i < $intend; $i++) $intend_string .= "\t";
		$pattern = "/:" . $key . "/";
		$values = array_intersect_key($input, array_flip(preg_grep($pattern, array_keys($input), $flags)));

		foreach ($values as $keyval => $value) {
			if ($value != "") {

				$shortLocale = explode(":", $keyval);

				$nodes .= $intend_string . "<" . $key . " locale=\"" . $locales[$shortLocale[0]] . "\">\r\n";

				$subvalues = explode(";", $value);
				foreach ($subvalues as $subvalue) {
					$nodes .= $intend_string . "\t<" . $key_singular . "><![CDATA[" . trim($subvalue) . "]]></" . $key_singular . ">\r\n";
				}

				$nodes .= $intend_string . "</" . $key . ">\r\n";
			}
		}

		return $nodes;
	}


	# Function for creating an array using the first row as keys
	function createArray($sheet)
	{
		$highestrow = $sheet->getHighestRow();
		$highestcolumn = $sheet->getHighestColumn();
		$columncount = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestcolumn);
		$headerRow = $sheet->rangeToArray('A1:' . $highestcolumn . "1");
		$header = $headerRow[0];
		array_unshift($header, "");
		unset($header[0]);
		$array = array();
		for ($row = 2; $row <= $highestrow; $row++) {
			$a = array();
			for ($column = 1; $column <= $columncount; $column++) {
				if (strpos($header[$column], "abstract") !== false) {
					if ($sheet->getCellByColumnAndRow($column, $row)->getValue() instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText) {
						$value = $sheet->getCellByColumnAndRow($column, $row)->getValue();
						$elements = $value->getRichTextElements();
						$cellData = "";
						foreach ($elements as $element) {
							if ($element instanceof \PhpOffice\PhpSpreadsheet\RichText\Run) {
								if ($element->getFont()->getBold()) {
									$cellData .= '<b>';
								} elseif ($element->getFont()->getSubScript()) {
									$cellData .= '<sub>';
								} elseif ($element->getFont()->getSuperScript()) {
									$cellData .= '<sup>';
								} elseif ($element->getFont()->getItalic()) {
									$cellData .= '<em>';
								}
							}
							// Convert UTF8 data to PCDATA
							$cellText = $element->getText();
							$cellData .= htmlspecialchars($cellText);
							if ($element instanceof \PhpOffice\PhpSpreadsheet\RichText\Run) {
								if ($element->getFont()->getBold()) {
									$cellData .= '</b>';
								} elseif ($element->getFont()->getSubScript()) {
									$cellData .= '</sub>';
								} elseif ($element->getFont()->getSuperScript()) {
									$cellData .= '</sup>';
								} elseif ($element->getFont()->getItalic()) {
									$cellData .= '</em>';
								}
							}
						}
						$a[$header[$column]] = $cellData;
					} else {
						$a[$header[$column]] = $sheet->getCellByColumnAndRow($column, $row)->getFormattedValue();
					}
				} else {
					$key = $header[$column];
					$a[$key] = $sheet->getCellByColumnAndRow($column, $row)->getFormattedValue();
				}
			}
			$array[$row] = $a;
		}

		return $array;
	}

	# Check the highest author number
	function countMaxAuthors($sheet)
	{
		$highestcolumn = $sheet->getHighestColumn();
		$headerRow = $sheet->rangeToArray('A1:' . $highestcolumn . "1");
		$header = $headerRow[0];
		$authorFirstnameValues = array();
		foreach ($header as $headerValue) {
			if ($headerValue && strpos($headerValue, "authorFirstname") !== false) {
				$authorFirstnameValues[] = (int) trim(str_replace("authorFirstname", "", $headerValue));
			}
		}
		return max($authorFirstnameValues);
	}

	# Check the highest file number
	function countMaxFiles($sheet)
	{
		$highestcolumn = $sheet->getHighestColumn();
		$headerRow = $sheet->rangeToArray('A1:' . $highestcolumn . "1");
		$header = $headerRow[0];
		$fileValues = array();
		foreach ($header as $headerValue) {
			if ($headerValue && strpos($headerValue, "fileLabel") !== false) {
				$fileValues[] = (int) trim(str_replace("fileLabel", "", $headerValue));
			}
		}
		return max($fileValues);
	}

	# Function for data validation
	function validateArticles($articles)
	{
		global $filesFolder;
		$errors = "";
		$articleRow = 0;

		foreach ($articles as $article) {

			$articleRow++;

			if (empty($article['issueYear'])) {
				$errors .= date('H:i:s') . " ERROR: Issue year missing for article " . $articleRow . EOL;
			}

			if (empty($article['issueDatePublished'])) {
				$errors .= date('H:i:s') . " ERROR: Issue publication date missing for article " . $articleRow . EOL;
			}

			if (empty($article['title'])) {
				$errors .= date('H:i:s') . " ERROR: article title missing for the given default locale for article " . $articleRow . EOL;
			}

			if (empty($article['sectionTitle'])) {
				$errors .= date('H:i:s') . " ERROR: section title missing for the given default locale for article " . $articleRow . EOL;
			}

			if (empty($article['sectionAbbrev'])) {
				$errors .= date('H:i:s') . " ERROR: section abbreviation missing for the given default locale for article " . $articleRow . EOL;
			}

			for ($i = 1; $i <= 200; $i++) {

				if (isset($article['file' . $i]) && $article['file' . $i] && !preg_match("@^https?://@", $article['file' . $i])) {

					$fileCheck = $this->filesFolder . $article['file' . $i];

					if (!file_exists($fileCheck))
						$errors .= date('H:i:s') . " ERROR: file " . $i . " missing " . $fileCheck . EOL;

					$fileLabelColumn = 'fileLabel' . $i;
					if (empty($fileLabelColumn)) {
						$errors .= date('H:i:s') . " ERROR: fileLabel " . $i . " missing for article " . $articleRow . EOL;
					}
					$fileLocaleColumns = 'fileLocale' . $i;
					if (empty($fileLocaleColumns)) {
						$errors .= date('H:i:s') . " ERROR: fileLocale " . $i . "  missingfor article " . $articleRow . EOL;
					}
				} else {
					break;
				}
			}
		}

		return $errors;
	}

	// get unique column keys statring with <name>
	function getUniqueKeys($articles, $name = NULL)
	{
		$uniqueKeys = array_unique(array_keys(array_merge(...$articles)));
		if (!$name) {
			return $uniqueKeys;
		}
		return array_filter($uniqueKeys, function ($key) use ($name) {
			return strpos($key, $name) === 0;
		});
	}
}

$app = new ConvertExcel2PKPNativeXML($argv);
