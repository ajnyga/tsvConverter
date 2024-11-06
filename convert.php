<?php

# Usage:
# php convert.php -x <xslx file> -f <files folder> [-v] [-l <default locale>]

# Hint: Use debugPrintXML($root) to write DOM to a file 'debug.xml'. $root should be a DOMDocument object.

// PHPExcel settings
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
define('EOL', (PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

require 'vendor/autoload.php';

class ConvertExcel2PKPNativeXML {

	// cli parsing
	private $opts;
	private $posArgs;
	private $fullFilesFolderPath;
	private $onlyValidate = false;
	private $xlsxFileName = 'articleData.xlsx';
	private $filesFolderName = 'files';

	// defaults
	private $defaultUploader = 'admin';
	private $defaultAuthor = ['givenname' => 'Editorial Board'];
	private $defaultLocale = 'en_US';
	private $defaultUserGroupRef = [
		'en_US' => 'Author',
		'de_DE' => 'Autor/in',
		'sv_SE' => 'F&#xF6;rfattare'
	];
	private $primaryContactId;

	// table parsing
	private $issueKeys;
	private $sectionKeys;
	private $articleKeys;
	private $authorKeys;
	private $locales = [
		'en' => 'en_US',
		'fi' => 'fi_FI',
		'sv' => 'sv_SE',
		'de' => 'de_DE',
		'ru' => 'ru_RU',
		'fr' => 'fr_FR',
		'no' => 'nb_NO',
		'da' => 'da_DK',
		'es' => 'es_ES',
	];	

	// xml generation
	private $articleElementOrder;
	private $publicationElementOrder;
	private $authorElementOrder;
	private $submissionFileElementOrder;
	private $issueElementOrder;
	private $issueIdentificationElementOrder;
	private $coverImageElementOrder;
	private $elementHasLocaleAttribute;
	
	// Constructor
	public function __construct($argv) {

		// pasre cli
		$rest_index = null;
		$shortOpts = "vl:x:f:";
		$longOpts = ['defaultLocale:', 'validate'];
		$this->opts = getopt($shortOpts, $longOpts, $rest_index);
		$this->posArgs = array_slice($argv, $rest_index);

		// Parse the INI configuration file
		foreach (parse_ini_file('config.ini', true) as $key => $value) {
			$this->{$key} = $value;
		}

		if (!$this->validateInput()) {
			echo date('H:i:s'), " Data validation failed!", EOL;
		}

		// Get the required order of elements from sample xml file
		$xsdFile = 'OJS_3.3_Native_Sample.xml';
		$dom = new DOMDocument;
		$dom->load($xsdFile);
		$this->articleElementOrder = $this->getChildElementsOrder($dom, 'article');
		$this->publicationElementOrder = $this->getChildElementsOrder($dom, 'pkppublication');
		array_splice($this->publicationElementOrder, 1, 0, 'doi'); // allow 'doi' to come directly after 'id'
		$this->authorElementOrder = $this->getChildElementsOrder($dom, 'author');
		$this->submissionFileElementOrder = $this->getChildElementsOrder($dom, 'submission_file');
		$this->issueElementOrder = $this->getChildElementsOrder($dom, 'issue');
		$this->issueIdentificationElementOrder = $this->getChildElementsOrder($dom, 'issue_identification');
		$this->coverImageElementOrder = $this->getChildElementsOrder($dom, 'cover');
		$this->elementHasLocaleAttribute = $this->hasLocaleAttribute($dom);

		// load data
		echo date('H:i:s'), " Creating a new PHPExcel object", EOL;
		$objReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($this->xlsxFileName);
		$objReader->setReadDataOnly(false);
		$objPhpSpreadsheet = $objReader->load($this->xlsxFileName);
		$sheet = $objPhpSpreadsheet->setActiveSheetIndex(0);

		echo date('H:i:s'), " Creating an array", EOL;
		$articles = $this->createArray($sheet);

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

		# Download galley files if fileName is empty or not provided and a gelleyDoi is available
		foreach ($articles as $index => &$article) {
			foreach ($article as $key => $value) {
				if (preg_match('/^fileName\d+$/', $key) && empty($value)) {
					$galleyDoiKey = str_replace('fileName', 'galleyDoi', $key);
					if (isset($article[$galleyDoiKey]) && !empty($article[$galleyDoiKey])) {
						$url = $article[$galleyDoiKey];
						$article[$key] = 'galleyFile' . ($index-array_key_first($articles)+1) . '.pdf';
						if (!file_exists($this->fullFilesFolderPath.$article[$key])) {
							$fileContent = file_get_contents($url);
							file_put_contents($this->fullFilesFolderPath.$article[$key], $fileContent);
							echo "Downloaded: ".$article[$key]." from $url\n";
						}
					}
				}
			}
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
		$this->authorKeys = $this->getUniqueKeys($articles, 'author');
		$this->articleKeys = array_diff($this->getUniqueKeys($articles), $this->issueKeys, $this->sectionKeys, $this->authorKeys);

		$issueIdentifications = [];
		$issueData = [];
		$articles = array_values($articles); // reindex array (keys used as article id later)
		foreach ($articles as $id => $article) {

			// identify issue and sections for each article by hash generated from issue and section fields

			// get issue data
			$issueIdentification = array_intersect_key($article, array_flip($this->issueKeys));
			$articleIssueHash = hash("sha256", implode(
				", ",
				[
					$issueIdentification['issueDatePublished'],
					$issueIdentification['issueVolume'],
					$issueIdentification['issueYear'],
					$issueIdentification['issueTitle']
				])
			);
			$issueIdentifications[$articleIssueHash] = $issueIdentification;
			foreach ($issueIdentification as $key => $value) {
				unset($article[$key]);
			}

			// get section data
			$sectionIdentification = array_intersect_key($article, array_flip($this->sectionKeys));
			// Sort the array alphabetically by the node value (required by native.xsd)
			ksort($sectionIdentification);
			foreach ($sectionIdentification as $key => $value) {
				unset($article[$key]);
			}

			// put all together			
			$issueData[$articleIssueHash]['issues'] = $issueIdentification;
			$issueData[$articleIssueHash]['sections']['section'] = $sectionIdentification;
			$issueData[$articleIssueHash]['articles'][$id] = $article;
			$issueData[$articleIssueHash]['articles'][$id]['sectionAbbrev'] = $sectionIdentification['sectionAbbrev'];
		}

		/* 
		* Create XML  
		* --------------------
		*/

		echo date('H:i:s'), " Starting XML output", EOL;

		$dom = new DOMDocument('1.0', 'UTF-8');
		$dom->formatOutput = true;

		[$issuesDOM, $pos] = $this->getOrCreateDOMElement($dom, 'issues', namespace: 'http://pkp.sfu.ca');
		$dom->appendChild($issuesDOM);

		// Create issue DOMs
		foreach ($issueData as $issueHash => $issueDataContent) {
			$issuesDOM = $this->processData($issuesDOM, $issueDataContent);
		}

		// reorder issue nodes
		foreach ($issuesDOM->childNodes as $issueDOM) {
			$issueDOM = $this->orderDOMNodes($issueDOM, $this->issueElementOrder);
		}

		$xpath = new DOMXPath($dom);
		$xpath->registerNamespace('xmlns', 'http://pkp.sfu.ca'); 
		$numberOfIssues = $xpath->query( "//xmlns:issue", $dom)->length;
		$numberOfSections = $xpath->query( "//section", $dom)->length;
		$numberOfArticles = $xpath->query( "//article", $dom)->length;
		$numberOfArticleGalleys = $xpath->query( "//xmlns:article_galley", $dom)->length;
		$numberOfEmbeds = $xpath->query( "//embed", $dom)->length;
		print_r("Info: Added $numberOfIssues issues, $numberOfSections sections, $numberOfArticles articles, $numberOfArticleGalleys gelleys and $numberOfEmbeds embedded elements to the XML file.\n");

		$dom->save(filename: $this->xlsxFileName.'.xml');
	}

	function validateInput() {
		// Check if the required parameter -x is set
		if (!isset($this->opts['x'])) {
			echo "Error: Required parameter -x <xlsx filename> is missing.\n";
			exit(1);
		}
		$this->xlsxFileName = $this->opts['x'];

		// Check if the required parameter -f is set
		if (!isset($this->opts['f'])) {
			echo "Error: Required parameter -f <files folder name> is missing.\n";
			exit(1);
		}
		$this->filesFolderName = $this->opts['f'];

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

		if (!file_exists($this->xlsxFileName)) {
			echo date('H:i:s') . " ERROR: Excel file does not exist" . EOL;
			die();
		}

		// Location of full text files
		$this->fullFilesFolderPath = dirname(__FILE__) . "/" . $this->filesFolderName . "/";

		if (!file_exists($this->fullFilesFolderPath)) {
			echo date('H:i:s') . " ERROR: given folder does not exist" . EOL;
			die();
		}

		return true;
	}

	function processData($dom, $data) {
		foreach ($data as $tagname => $content) {
			if (strlen($tagname) > 0) // to reject any blank lines in the excel sheet
			switch ($tagname) {
				case 'sections':
					[$issueDOM, $pos] = $this->getOrCreateDOMElement($dom->ownerDocument, 'issue');

					[$sectionsDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'sections');
					$issueDOM->appendChild($sectionsDOM);

					[$sectionDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'section');
					$sectionsDOM->appendChild($sectionDOM);

					$sectionDOM = $this->processData($sectionDOM, $data['sections']['section']);	
					break;
				case 'issues':
					[$issueDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'issue', 'http://pkp.sfu.ca');
					$dom->appendChild($issueDOM);

					$issueData = $data[$tagname];
					$issueData = $this->stripColumnPrefix($issueData, 'issue');

					$issueIdentificationData = [];
					foreach ($this->issueIdentificationElementOrder as $field) {
						foreach (array_keys($issueData) as $key) {
							if (str_ends_with($key,$field)) {
								$issueIdentificationData[$key] = $issueData[$key];
								unset($issueData[$key]);
							}
						}
					}
					
					$issueData['issue_identification'] = $issueIdentificationData;

					$issueData = $this->sortArrayElementsByKey($issueData, $this->issueElementOrder);

					$issueDOM = $this->processData($issueDOM, $issueData);

					$issueDOM->setAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');
					$issueDOM->setAttribute('xsi:schemaLocation', 'http://pkp.sfu.ca native.xsd');
					$issueDOM->setAttribute('published', '1');
					$issueDOM->setAttribute('current', '0');
					break;
				case 'issue_identification':
					[$issuesIdentificationDOM, $pos] = $this->createDOMElement($dom->ownerDocument, $tagname);
					$dom->appendChild($issuesIdentificationDOM );

					$issuesIdentificationDOM = $this->processData($issuesIdentificationDOM, $content);
					break;
				case 'sectionTitle':
				case 'sectionAbbrev':
					// according to native.xsd abbrev and title need to be in (probably) alphabetic order (see ksort above)
					$xmlTagName = strtolower(str_replace('section', '', $tagname));
					$dom = $this->processData($dom, [$xmlTagName => $content]);
					$dom->setAttribute('ref', $data['sectionAbbrev']);
					$dom->setAttribute('seq', isset($data['sectionSeq']) ? $data['sectionSeq'] : "0");
					break;
				case 'datePublished':
					[$issueDOM, $pos] = $this->getOrCreateDOMElement($dom->ownerDocument, 'issue');
					$element = $dom->ownerDocument->createElement('date_published', $content);
					$issueDOM->appendChild($element);
					$element = $dom->ownerDocument->createElement('last_modified', $content);
					$issueDOM->appendChild($element);
					break;
				case 'articles':
					[$articlesDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'articles', 'http://pkp.sfu.ca');
					[$issueDOM, $pos] = $this->getOrCreateDOMElement($dom->ownerDocument, 'issue');
					$issueDOM->appendChild($articlesDOM);

					foreach ($content as $articleId => $article) {

						# Article
						echo date('H:i:s'), " Adding article: ", $article['title'], EOL;

						[$articleDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'article');
						$articlesDOM->appendChild($articleDOM);

						$articleDOM->setAttribute('stage', 'production');
						$issueDatePublished = $dom->getElementsByTagName('date_published')[0]->textContent;
						$articleDOM->setAttribute('date_submitted', $issueDatePublished);
						$articleDOM->setAttribute('status', '3');
						$articleDOM->setAttribute('submission_progress', '0');
						$articleDOM = $this->processData($articleDOM, ['id' => [
								'type'=> 'internal',
								'id' => $articleId+1
							]]);

						// get file data 
						$fileKeys = $this->getUniqueKeys([$article], 'file');
						$fileData = [];
						foreach ($fileKeys as $key) {
							preg_match('/^(.*?)(\d+)$/', str_replace('file','',$key), $matches);
							$elementName = strtolower($matches[1]);
							$id = $matches[2];
							if (strlen($article[$key]) > 0) {
								$fileData[$id][$elementName] = $article[$key];
							}
							unset($article[$key]);
						}

						$articleDOM = $this->processData($articleDOM, [
							'submission_file' => $fileData
						]);
							
						$articleDOM = $this->processData($articleDOM, [
							'publication' => $article
						]);
					}
	
					break;
				case 'publication':
					[$publicationDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'publication');
					$dom->appendChild($publicationDOM);
					$publicationId = $pos;

					$dom->setAttribute('current_publication_id', $publicationId);

					$publicationDOM->setAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');
					$publicationDOM->setAttribute('xsi:schemaLocation', 'http://pkp.sfu.ca native.xsd');
					
					# Check if language has an alternative default locale
					# If it does, use the locale in all fields
					$articleLocale = $this->defaultLocale;
					if (!empty($article['language'])) {
						$articleLocale = $this->locales[trim($content['language'])];
					}
					unset($content['language']);

					$publicationDOM->setAttribute('locale', $articleLocale);
					$publicationDOM->setAttribute('version', "1");
					$publicationDOM->setAttribute('status', "3");
					$publicationDOM->setAttribute('access_status', "0");
					$publicationDOM->setAttribute('url_path', "");

					[$element, $pos] = $this->getOrCreateDOMElement($dom->ownerDocument,'issue');
					$datePublishedList = $element->getElementsByTagName('date_published');
					if ($datePublishedList->length > 0) {
						$publicationDOM->setAttribute('date_published', $datePublishedList[0]->textContent);
					}
					
					$publicationDOM->setAttribute('section_ref', $content['sectionAbbrev']);
					unset($content['sectionAbbrev']);
					if (isset($content['articleSeq'])) {
						$publicationDOM->setAttribute('seq', $content['articleSeq']);
						unset($content['articleSeq']);
					}

					$publicationDOM = $this->processData($publicationDOM,  [
						'id' => [
							'type'=> 'internal',
							'id' => $publicationId
						]]);

					$content = $this->stripColumnPrefix($content, 'article');

					//  get the author data (we need to process it later)
					$authorKeys = $this->getUniqueKeys([$content], 'author');
					$content['authors'] = [];
					foreach ($authorKeys as $key) {
						preg_match('/^(.*?)(\d+)$/', str_replace('author','',$key), $matches);
						$elementName = strtolower($matches[1]);
						$id = $matches[2];
						if (strlen($content[$key]) > 0) {
							$content['authors'][$id][$elementName] = $content[$key];
						}
						unset($content[$key]);
					}
					foreach ($content['authors'] as $id => $authorData) {
						// create required fields if not provided
						$missingKeys = array_diff(
							['givenname','familyname','affiliation','country','email'],
							array_keys($authorData)
						);
						// set missing values
						foreach ($missingKeys as $key) {
							if ($key == 'givenname') {
								$authorData[$key] = $this->defaultAuthor;
							} else {
								$authorData[$key] = "";
							}
						}
						// sort elements according to required field order
						$content['authors'][$id] = $this->sortArrayElementsByKey($authorData, $this->authorElementOrder);
					}

					// get galley data 
					$galleyKeys = $this->getUniqueKeys([$content], 'galley');
					$galleyData = [];
					foreach ($galleyKeys as $key) {
						preg_match('/^(.*?)(\d+)$/', str_replace('galley','',$key), $matches);
						$elementName = strtolower($matches[1]);
						$id = $matches[2];
						if (strlen($content[$key]) > 0) {
							$galleyData[$id][$elementName] = $content[$key];
						}
						unset($content[$key]);
					}
					$content['article_galley'] = $galleyData;

					$content = $this->sortArrayElementsByKey($content, $this->publicationElementOrder);

					// process data
					$publicationDOM = $this->processData($publicationDOM, $content);

					$publicationDOM->setAttribute('primary_contact_id', $this->primaryContactId);
					break;
				case 'authors':
					[$authorsDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'authors', 'http://pkp.sfu.ca');
					$dom->appendChild($authorsDOM);

					$i = 0;
					foreach ($content as $authorId => $author) {
							
						[$authorDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'author');
						$authorsDOM->appendChild($authorDOM);
				
						$authorDOM->setAttribute('include_in_browse', 'true');
						$authorDOM->setAttribute('user_group_ref', $this->defaultUserGroupRef[$this->defaultLocale]);
						$authorDOM->setAttribute('seq', $authorId);
						$authorDOM->setAttribute('id', $pos + 1 + $i);

						$this->primaryContactId = $pos + 1;
						if (isset($data['primaryContactId'])) {
							if ($data['primaryContactId'] == $authorId) {
								$this->primaryContactId = $pos + 1 + $i;
							}
							unset($data['primaryContactId']);
						}

						$authorDOM = $this->processData($authorDOM, $author);
						$i++;
					}
					break;
				case 'article_galley':
					foreach ($content as $id => $galleyData) {
						[$articleGalleysDOM, $pos] = $this->createDOMElement($dom->ownerDocument, 'article_galley', 'http://pkp.sfu.ca');
						$dom->appendChild($articleGalleysDOM);
						
						$articleGalleysDOM->setAttribute('locale', $this->locales[$galleyData['locale']]);
						$articleGalleysDOM->setAttribute('approved', "false");

						$articleGalleysDOM = $this->processData($articleGalleysDOM, ['name' => $galleyData['label']]);
						$articleGalleysDOM = $this->processData($articleGalleysDOM, ['seq' => $id-1]);

						[$fileRef, $pos] = $this->createDOMElement($dom->ownerDocument, 'submission_file_ref');
						$articleGalleysDOM->appendChild($fileRef);

						$fileRef->setAttribute('id', $pos+1);
					}
					break;
				case 'id':
					switch ($content['type']) {
						case 'internal':
							$id = $dom->ownerDocument->createElement('id', $content['id']);
							$id->setAttribute('type', 'internal');
							$id->setAttribute('advice', 'ignore');
							break;
					}
					$dom->appendChild($id);
					break;
				case 'doi':
					$id = $dom->ownerDocument->createElement('id', $content);
					$dom->appendChild($id);

					$id->setAttribute('type', 'doi');
					$id->setAttribute('advice', 'update');
					break;
				case str_ends_with($tagname, 'keywords'):
				case str_ends_with($tagname, 'disciplines'):
				case str_ends_with($tagname, 'subjects'):
					if (strlen($content) > 0) {
						[$locale, $xmlTagName] = $this->splitLocaleTagName($tagname);
						[$elementsDOM, $pos] = $this->createDOMElement($dom->ownerDocument, $xmlTagName);
						$dom->appendChild($elementsDOM);
	
						$elementsDOM->setAttribute('locale', $locale);
						
						foreach (explode(';', $content) as $element) {
							$elementDOM = $dom->ownerDocument->createElement(rtrim($xmlTagName, "s"), $element);
							$elementsDOM->appendChild($elementDOM);
						}
					}					
					break;
				case 'submission_file':
					foreach ($content as $id => $submissionFileData) {
						[$subFileDOM, $pos] = $this->createDOMElement($dom->ownerDocument, $tagname, 'http://pkp.sfu.ca');
						$dom->appendChild($subFileDOM);

						$subFileDOM->setAttribute('stage', 'proof');
						$subFileDOM->setAttribute('id', $pos+1);
						$subFileDOM->setAttribute('file_id', $pos+1);
						$subFileDOM->setAttribute('uploader', $this->defaultUploader);
						$subFileDOM->setAttribute('genre', $submissionFileData['genre']);
						unset($submissionFileData['genre']);

						$submissionFileData = $this->sortArrayElementsByKey($submissionFileData, $this->submissionFileElementOrder);
	
						$subFileDOM = $this->processData($subFileDOM, $submissionFileData);
						
						$filePath = $this->fullFilesFolderPath . $submissionFileData['name'];
						if (file_exists($filePath)) {
							$size = filesize($filePath);
							echo date('H:i:s') . " Adding file " . $filePath . EOL;
							$file = $dom->ownerDocument->createElement('file');
							$subFileDOM->appendChild($file);

							$file->setAttribute('id', $pos+1);
							$file->setAttribute('filesize', $size);
							$file->setAttribute('extension', pathinfo($submissionFileData['name'], PATHINFO_EXTENSION));

							$embed = $dom->ownerDocument->createElement('embed', base64_encode(file_get_contents($filePath)));
							$embed->setAttribute('encoding','base64');
							$file->appendChild($embed);
						} else {
							echo date('H:i:s') . " WARNING: file " . $filePath . " not found !" . EOL;
						}
					}
					break;
				case str_ends_with($tagname, 'coverImage'):
				case str_ends_with($tagname, 'coverImageAltText'):
					if (strlen($content) > 0) {

						[$locale, $xmlTagName] = $this->splitLocaleTagName($tagname);
						[$coverDOM, $pos] = $this->getOrCreateDOMElement($dom, 'cover', $locale);
						if ($coverDOM->childElementCount == 0) {
							[$coversDOM, $pos] = $this->getOrCreateDOMElement($dom, 'covers', namespace: 'http://pkp.sfu.ca');
							$coversDOM->appendChild($coverDOM);
							$dom->appendChild($coversDOM);
							$coverDOM->setAttribute('locale', $locale);
						}

						if ($xmlTagName == 'coverImageAltText') {
							$node = $coverDOM->getElementsByTagName('cover_image_alt_text')[0];
							if ($node) {
								$coverDOM->removeChild($node);
							}
							$coverDOM = $this->processData($coverDOM, ['cover_image_alt_text' => $content]);
						} else {
							$coverDOM = $this->processData($coverDOM, ['cover_image' => $content]);
							$filePath = $this->fullFilesFolderPath . $content;
							if (file_exists($filePath)) {
								$embed = $dom->ownerDocument->createElement('embed', base64_encode(file_get_contents($filePath)));
								$embed->setAttribute('encoding','base64');
								$coverDOM->appendChild($embed);
							}
						}

						//reorder nodes and set default alt text
						$childNodes = ['cover_image_alt_text' => $dom->ownerDocument->createElement('cover_image_alt_text',"")];
						foreach ($coverDOM->childNodes as $child) {
							$childNodes[$child->tagName] = $child;
						}
						$childNodes = $this->sortArrayElementsByKey($childNodes, $this->coverImageElementOrder);
						foreach ($childNodes as $child) {
							$coverDOM->appendChild($child);
						}
					}
					break;
				default:
					// here we handle all text nodes
					if (strlen($tagname) > 0) {
						switch ($tagname) {
							case 'primaryContactId':
								// fields that hold attributes don't create a tag
								break;
							case in_array($tagname, $this->elementHasLocaleAttribute):
							case (strpos($tagname, ':') === 2):
								// elements with locale attribute
								[$locale, $tagname] = $this->splitLocaleTagName($tagname);
								$element = $dom->ownerDocument->createElement($tagname);
								$element->appendChild($dom->ownerDocument->createTextNode($content));
								if ($locale) {
									$element->setAttribute('locale', $locale);
								}
								$dom->appendChild($element);
								break;
							case 'copyrightYear':
								if (strlen($content) == 0) {
									break;
								}
							default:
								// elements without locale attribute
								$element = $dom->ownerDocument->createElement($tagname);
								$element->appendChild($dom->ownerDocument->createTextNode($content));
								$dom->appendChild($element);
								break;
						}
					}
					break;
			}
		}
		return $dom;
	}

	/* 
	* Helpers 
	* -----------
	*/

	// sort elements according to required field order
	function sortArrayElementsByKey(array $dataArray, array $fieldOrder) {

		$orderedArray = [];
		foreach ($fieldOrder as $xmlTagName) {
			if (isset($dataArray[$xmlTagName])) {
				$orderedArray[$xmlTagName] =  $dataArray[$xmlTagName];
				unset($dataArray[$xmlTagName]);
			}
			foreach (array_flip($this->locales) as $locale) {
				$localeKey = $locale.':'.$xmlTagName;
				if (array_key_exists($localeKey, $dataArray)) {
					$orderedArray[$localeKey] = $dataArray[$localeKey];
					unset($dataArray[$localeKey]);
				}
			}
		}
		$orderedArray = array_merge($orderedArray, $dataArray);
		return $orderedArray;
	}

	// extract locale value from tag name
	function splitLocaleTagName($tagname, $locale = NULL) {
		// Is there a valid locale specified?
		if (strpos($tagname, ":") !== false) {
			$locale = $this->locales[explode(':',$tagname)[0]];
			if (!$locale) {
				$locale = $this->defaultLocale;
			} else {
				$tagname = explode(':',$tagname)[1];
			}
			return [$locale, $tagname];
		} else {
			$locale = $this->defaultLocale;
		}
		return [$locale, $tagname];
	}

	function createDOMElement($root, $tagname, $namespace = NULL) {
		echo date('H:i:s') . " Creating element " . $tagname . EOL;
		if ($namespace) {
			$targetDOM = $root->createElementNS($namespace, $tagname);
			$targetDOM->setAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');
			$targetDOM->setAttribute('xsi:schemaLocation', 'http://pkp.sfu.ca native.xsd');
		} else {
			$targetDOM = $root->createElement($tagname);
		}
		
		return [$targetDOM, $root->getElementsByTagname($tagname)->length];
	}

	// Try to get the last DOM element with the given tag name. If none exists create a new one
	function getOrCreateDOMElement($dom, $tagname, $locale = NULL, $namespace = NULL) {
		// try to get the requested element
		if (get_class($dom) !== "DOMDocument") {
			$root = $dom->ownerDocument;
		} else {
			$root = $dom;
		}
		$targetDOM = $dom->getElementsByTagName($tagname);
		if (($targetDOM->length > 0) && $locale) {
			$xpath = new DOMXPath($root);
			$targetDOM = $xpath->query(
				expression: "//".$tagname."[@locale='$locale']",
				contextNode: $dom
			); 
		}
		
		// create element if not found
		if (!isset($targetDOM) || $targetDOM->length == 0) {
			[$targetDOM, $pos] = $this->createDOMElement($root, $tagname, $namespace);
			if (!get_class($dom) === "DOMDocument") {
				$dom->lastChild->appendChild($targetDOM);
			} else {
				$dom->appendChild($targetDOM);
			}
			$elementPosition = $pos;
		} else {
			$elementPosition = $targetDOM->length;
			$targetDOM = $targetDOM->item($targetDOM->length - 1);
		}

		return [$targetDOM, $elementPosition];
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

	# Function for data validation
	function validateArticles($articles)
	{
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

					$fileCheck = $this->fullFilesFolderPath . $article['file' . $i];

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

	// get unique column keys starting with <name>
	function getUniqueKeys($articles, $name = NULL)
	{
		$uniqueKeys = array_unique(array_keys(array_merge(...$articles)));
		if (!$name) {
			return $uniqueKeys;
		}
		return array_filter($uniqueKeys, function ($key) use ($name) {
			return ((strpos($key, $name) === 0) || (strpos($key, $name) === 3)); // does the name occur at the beginning or after a locale code?
		});
	}

	// Function to find the specific element and get its child elements
	function getChildElementsOrder($node, $elementName) {
		$order = [];

		// Check if the node is the target element
		if ($node->nodeType === XML_ELEMENT_NODE && $node->nodeName === $elementName) {
			if ($node->hasChildNodes()) {
				foreach ($node->childNodes as $child) {
					if ($child->nodeType === XML_ELEMENT_NODE) {
						$order[] = $child->nodeName; // Add the element name to the order
					}
				}
			}
		}

		// Recursively search for the target element in child nodes
		if ($node->hasChildNodes()) {
			foreach ($node->childNodes as $child) {
				$order = array_merge($order, $this->getChildElementsOrder($child, $elementName));
			}
		}

		return array_unique($order);
	}

	// return an array with element names that have the "locale" attribute
	function hasLocaleAttribute($dom) {

		$elementsWithLocale = [];

		$elements = $dom->getElementsByTagName('*');
		foreach ($elements as $element) {
			if ($element->hasAttribute('locale')) {
				$elementsWithLocale[] = $element->nodeName;
			}
		}
		return $elementsWithLocale;
	}

	// strip 'article' prefix from key names
	function stripColumnPrefix(array $data, string $prefix) {
		foreach ($data as $key => $value) {
			// Check if the key starts with 'article'
			if ((strpos($key, $prefix) === 0) || (strpos($key, $prefix) === 3)) {
				// Remove 'article' from the key
				if (strpos($key, $prefix) === 3) {
					// there is a locale descriptor we need to consider
					$newKey = str_replace($prefix, '', $key);
					$keyParts = explode(':', $newKey);
					$keyParts[1]= lcfirst($keyParts[1]);
					$key = implode(':', $keyParts);
					$newKey = str_replace($prefix, '', $key);
				} else {
					$newKey = lcfirst(str_replace($prefix, '', $key));
				}
				$newArray[$newKey] = $value; // Assign the value to the new key
			} else {
				$newArray[$key] = $value; // Keep the original key-value pair
			}
		}
		return $newArray;
	}

	function orderDOMNodes($dom, $order) {
		$childNodes = [];
		foreach ($dom->childNodes as $child) {
			$childNodes[$child->tagName] = $child;
		}
		$childNodes = $this->sortArrayElementsByKey($childNodes, $order);
		foreach ($childNodes as $child) {
			$dom->appendChild($child);
		}
		return $dom;
	}

}

$app = new ConvertExcel2PKPNativeXML($argv);

function debugPrintXML($root) {
	$root->save('debug.xml');
}