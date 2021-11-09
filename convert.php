<?php

/* 
 * Settings 
 * ---------
 */

$fileName = '';
$onlyValidate = 0;
$files = "files";

// Get arguments from command line
if (isset($argv[1])) {
	$fileName = $argv[1];
}
if (isset($argv[2])) {
	$files = $argv[2];
}
if (isset($argv[3]) && $argv[3] == '-v') {
	$onlyValidate = 1;
}


// The default locale. For alternative locales use language field. For additional locales use locale:fieldName.
$defaultLocale = 'fi_FI';

// The uploader account name
$uploader = "admin";

// Default author name. If no author is given for an article, this name is used instead.
$defaultAuthor['givenname'] = "Editorial Board";

// Location of full text files
$filesFolder = dirname(__FILE__) . "/". $files ."/";

// Possible locales
$locales = array(
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

// PHPExcel settings
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/Helsinki');
define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/* 
 * Check that a file and a folder exists
 * ------------------------------------
 */
if (!file_exists($fileName)) {
	echo date('H:i:s') . " ERROR: given file does not exist" . EOL;
	die();
}

if (!file_exists($filesFolder)) {
	echo date('H:i:s') . " ERROR: given folder does not exist" . EOL;
	die();
}

/* 
 * Load Excel data to an array
 * ------------------------------------
 */
echo date('H:i:s') , " Creating a new PHPExcel object" , EOL;

$objReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($fileName);
$objReader->setReadDataOnly(false);
$objPhpSpreadsheet = $objReader->load($fileName);
$sheet = $objPhpSpreadsheet->setActiveSheetIndex(0);

echo date('H:i:s') , " Creating an array" , EOL;

$articles = createArray($sheet);
$maxAuthors = countMaxAuthors($sheet);
$maxFiles = countMaxFiles($sheet);

/* 
 * Data validation   
 * -----------
 */

echo date('H:i:s') , " Validating data" , EOL;

$errors = validateArticles($articles);
if ($errors != ""){
	echo $errors, EOL;
	die();	
}

# If only validation is selected, exit
if ($onlyValidate == 1){
	echo date('H:i:s') , " Validation complete " , EOL;
	die();
}


/* 
 * Prepare data for output
 * ----------------------------------------
 */

echo date('H:i:s') , " Preparing data for output" , EOL;

# Save section data
foreach ($articles as $article){
	$sections[$article['issueDatepublished']][$article['sectionAbbrev']] = $article['sectionTitle'];
}


/* 
 * Create XML  
 * --------------------
 */

echo date('H:i:s') , " Starting XML output" , EOL;
$currentIssueDatepublished = null;	
$currentYear = null;
$fileId = 1;


	foreach ($articles as $key => $article){
	
	# Issue :: if issueDatepublished has changed, start a new issue
	if ($currentIssueDatepublished != $article['issueDatepublished']){
		
		$newYear = date('Y', strtotime($article['issueDatepublished']));

		# close old issue if one exists
		if ($currentIssueDatepublished != null){
			fwrite ($xmlfile,"\t\t</articles>\r\n");
			fwrite ($xmlfile,"\t</issue>\r\n\r\n");
		}
		
		
		# Start a new XML file if year changes
		if ($newYear != $currentYear){

			if ($currentYear != null){
				echo date('H:i:s') , " Closing XML file" , EOL;
				fwrite ($xmlfile,"</issues>\r\n\r\n");
			}
			
			echo date('H:i:s') , " Creating a new XML file ", $newYear, ".xml" , EOL;
			
			$xmlfile = fopen ($newYear.'.xml','w');
			fwrite ($xmlfile,"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n");
			fwrite ($xmlfile,"<issues xmlns=\"http://pkp.sfu.ca\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"http://pkp.sfu.ca native.xsd\">\r\n");
		}
		
		fwrite ($xmlfile,"\t<issue xmlns=\"http://pkp.sfu.ca\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" published=\"1\" xsi:schemaLocation=\"http://pkp.sfu.ca native.xsd\">\r\n\r\n");
		
		echo date('H:i:s') , " Adding issue with publishing date ", $article['issueDatepublished'] , EOL;

		# Issue description
		if (!empty($article['issueDescription']))
			fwrite ($xmlfile,"\t\t<description locale=\"".$defaultLocale."\"><![CDATA[".$article['issueDescription']."]]></description>\r\n");

		# Issue identification
		fwrite ($xmlfile,"\t\t<issue_identification>\r\n");
		
		if (!empty($article['issueVolume']))
			fwrite ($xmlfile,"\t\t\t<volume><![CDATA[".$article['issueVolume']."]]></volume>\r\n");	
		if (!empty($article['issueNumber']))
			fwrite ($xmlfile,"\t\t\t<number><![CDATA[".$article['issueNumber']."]]></number>\r\n");			
		fwrite ($xmlfile,"\t\t\t<year><![CDATA[".$article['issueYear']."]]></year>\r\n");
		
		if (!empty($article['issueTitle'])){
			fwrite ($xmlfile,"\t\t\t<title><![CDATA[".$article['issueTitle']."]]></title>\r\n");
		}
		# Add alternative localisations for the issue title
		fwrite ($xmlfile, searchLocalisations('issueTitle', $article, 3));
		
		fwrite ($xmlfile,"\t\t</issue_identification>\r\n\r\n");
		
		fwrite ($xmlfile,"\t\t<date_published><![CDATA[".$article['issueDatepublished']."]]></date_published>\r\n\r\n");
		
		# Sections
		fwrite ($xmlfile,"\t\t<sections>\r\n");
		    
			foreach ($sections[$article['issueDatepublished']] as $sectionAbbrev => $sectionTitle){
				fwrite ($xmlfile,"\t\t\t<section ref=\"".htmlentities($sectionAbbrev, ENT_XML1)."\">\r\n");
				fwrite ($xmlfile,"\t\t\t\t<abbrev locale=\"".$defaultLocale."\">".htmlentities($sectionAbbrev, ENT_XML1)."</abbrev>\r\n");
				fwrite ($xmlfile,"\t\t\t\t<title locale=\"".$defaultLocale."\"><![CDATA[".$sectionTitle."]]></title>\r\n");
				fwrite ($xmlfile, searchLocalisations('sectionTitle', $article, 3));
				fwrite ($xmlfile,"\t\t\t</section>\r\n");
			}
		
		fwrite ($xmlfile,"\t\t</sections>\r\n\r\n");
		
		# Issue galleys
		fwrite ($xmlfile,"\t\t<issue_galleys xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"http://pkp.sfu.ca native.xsd\"/>\r\n\r\n");
		
		# Start articles output
		fwrite ($xmlfile,"\t\t<articles xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"http://pkp.sfu.ca native.xsd\">\r\n\r\n");
		
		$currentIssueDatepublished = $article['issueDatepublished'];
		$currentYear = $newYear;
			
	}
	
	# Article
	echo date('H:i:s') , " Adding article: ", $article['title'] , EOL;

	# Check if language has an alternative default locale
	# If it does, use the locale in all fields
	$articleLocale = $defaultLocale;
	if (!empty($article['language'])){
		$articleLocale = $locales[trim($article['language'])];
	}
	
	fwrite ($xmlfile,"\t\t<article xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" locale=\"".$articleLocale."\" stage=\"production\" date_submitted=\"".$article['issueDatepublished']."\" date_published=\"".$article['issueDatepublished']."\" section_ref=\"".htmlentities($article['sectionAbbrev'], ENT_XML1)."\">\r\n\r\n");

		# DOI
		if (!empty($article['doi'])){	
			fwrite ($xmlfile,"\t\t\t<id type=\"doi\" advice=\"update\"><![CDATA[".$article['doi']."]]></id>\r\n");
		}
	
		# title, prefix, subtitle, abstract
		fwrite ($xmlfile,"\t\t\t<title locale=\"".$articleLocale."\"><![CDATA[".$article['title']."]]></title>\r\n");
		fwrite ($xmlfile, searchLocalisations('title', $article, 3));

		if (!empty($article['prefix'])){
			fwrite ($xmlfile,"\t\t\t<prefix locale=\"".$articleLocale."\"><![CDATA[".$article['prefix']."]]></prefix>\r\n");
		}
		fwrite ($xmlfile, searchLocalisations('prefix', $article, 3));
		
		if (!empty($article['subtitle'])){	
			fwrite ($xmlfile,"\t\t\t<subtitle locale=\"".$articleLocale."\"><![CDATA[".$article['subtitle']."]]></subtitle>\r\n");
		}
		fwrite ($xmlfile, searchLocalisations('subtitle', $article, 3));
		
		if (!empty($article['abstract'])){
			fwrite ($xmlfile,"\t\t\t<abstract locale=\"".$articleLocale."\"><![CDATA[".nl2br($article['abstract'])."]]></abstract>\r\n\r\n");
		}
		fwrite ($xmlfile, searchLocalisations('abstract', $article, 3));

		if (!empty($article['articleLicenseUrl'])) {	
			fwrite ($xmlfile,"\t\t\t<licenseUrl><![CDATA[".$article['articleLicenseUrl']."]]></licenseUrl>\r\n");	
		}	
		if (!empty($article['articleCopyrightHolder'])) {	
			fwrite ($xmlfile,"\t\t\t<copyrightHolder locale=\"".$articleLocale."\"><![CDATA[".$article['articleCopyrightHolder']."]]></copyrightHolder>\r\n");	
		}	
		if (!empty($article['articleCopyrightYear'])) {	
			fwrite ($xmlfile,"\t\t\t<copyrightYear><![CDATA[".$article['articleCopyrightYear']."]]></copyrightYear>\r\n");	
		}

		# Keywords
		if (!empty($article['keywords'])){
			if (trim($article['keywords']) != ""){
				fwrite ($xmlfile,"\t\t\t<keywords locale=\"".$articleLocale."\">\r\n");
				$keywords = explode(";", $article['keywords']);
				foreach ($keywords as $keyword){
					fwrite ($xmlfile,"\t\t\t\t<keyword><![CDATA[".trim($keyword)."]]></keyword>\r\n");	
				}
				fwrite ($xmlfile,"\t\t\t</keywords>\r\n");
			}
			fwrite ($xmlfile, searchTaxonomyLocalisations('keywords', 'keyword', $article, 3));
		}


		# Disciplines
		if (!empty($article['disciplines'])){
			if (trim($article['disciplines']) != "") {
				fwrite ($xmlfile,"\t\t\t<disciplines locale=\"".$articleLocale."\">\r\n");
				$disciplines = explode(";", $article['disciplines']);
				foreach ($disciplines as $discipline){
					fwrite ($xmlfile,"\t\t\t\t<discipline><![CDATA[".trim($discipline)."]]></discipline>\r\n");	
				}
				fwrite ($xmlfile,"\t\t\t</disciplines>\r\n");
			}
			fwrite ($xmlfile, searchTaxonomyLocalisations('disciplines', 'disciplin', $article, 3));
		}
		
		# TODO: add support for subjects, supporting agencies
		/*
		<agencies locale="fi_FI">
			<agency></agency>
		</agencies>
		<subjects locale="fi_FI">
			<subject></subject>
		</subjects>
		*/

		# Authors
		fwrite ($xmlfile,"\t\t\t<authors xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"http://pkp.sfu.ca native.xsd\">\r\n");
		
		for ($i = 1; $i <= $maxAuthors; $i++) {
			
			if ($article['authorFirstname'.$i]) {
				
				if ($i == 1)
					fwrite ($xmlfile,"\t\t\t\t<author primary_contact=\"true\" include_in_browse=\"true\" user_group_ref=\"Author\">\r\n");
				else
					fwrite ($xmlfile,"\t\t\t\t<author include_in_browse=\"true\" user_group_ref=\"Author\">\r\n");
				
				fwrite ($xmlfile,"\t\t\t\t\t<givenname><![CDATA[".$article['authorFirstname'.$i].(!empty($article['authorMiddlename'.$i]) ? ' '.$article['authorMiddlename'.$i] : '')."]]></givenname>\r\n");
				if (!empty($article['authorLastname'.$i])){
					fwrite ($xmlfile,"\t\t\t\t\t<familyname><![CDATA[".$article['authorLastname'.$i]."]]></familyname>\r\n");
				}

				if (!empty($article['authorAffiliation'.$i])){
					fwrite ($xmlfile,"\t\t\t\t\t<affiliation locale=\"".$articleLocale."\"><![CDATA[".$article['authorAffiliation'.$i]."]]></affiliation>\r\n");
				}
				fwrite ($xmlfile, searchLocalisations('authorAffiliation'.$i, $article, 5, 'affiliation'));

				if (!empty($article['country'.$i])){
					fwrite ($xmlfile,"\t\t\t\t\t<country><![CDATA[".$article['country'.$i]."]]></country>\r\n");
				}

				if (!empty($article['authorEmail'.$i])){
					fwrite ($xmlfile,"\t\t\t\t\t<email>".$article['authorEmail'.$i]."</email>\r\n");
				}
				else{
					fwrite ($xmlfile,"\t\t\t\t\t<email><![CDATA[]]></email>\r\n");
				}

				if (!empty($article['orcid'.$i])){
					fwrite ($xmlfile,"\t\t\t\t\t<orcid>".$article['orcid'.$i]."</orcid>\r\n");
				}
				if (!empty($article['authorBio'.$i])){
					fwrite ($xmlfile,"\t\t\t\t\t<biography locale=\"".$articleLocale."\"><![CDATA[".$article['authorBio'.$i]."]]></biography>\r\n");
				}
				
				fwrite ($xmlfile,"\t\t\t\t</author>\r\n");

				
			}

		}

		# If no authors are given, use default author name
		if (!$article['authorFirstname1']){
				fwrite ($xmlfile,"\t\t\t\t<author primary_contact=\"true\" user_group_ref=\"Author\">\r\n");
				fwrite ($xmlfile,"\t\t\t\t\t<givenname><![CDATA[".$defaultAuthor['givenname']."]]></givenname>\r\n");
				fwrite ($xmlfile,"\t\t\t\t\t<email><![CDATA[]]></email>\r\n");
				fwrite ($xmlfile,"\t\t\t\t</author>\r\n");
		}

		fwrite ($xmlfile,"\t\t\t</authors>\r\n\r\n");
	
	
		# Submission files
		unset($galleys);
		$fileSeq = 0;
		
		for ($i = 1; $i <= $maxFiles; $i++) {

			if (empty($article['fileLocale'.$i])) {
				$fileLocale = $articleLocale;
			} else {
				$fileLocale = $locales[trim($article['fileLocale'.$i])];
			}

			if (!preg_match("@^https?://@", $article['file'.$i]) && $article['file'.$i] != "") {
					
				$file = $filesFolder.$article['file'.$i];
				$fileSize = filesize($file);				
				if(function_exists('mime_content_type')){
					$fileType = mime_content_type($file);
				}
				elseif(function_exists('finfo_open')){
					$fileinfo = new finfo();
					$fileType = $fileinfo->file($file, FILEINFO_MIME_TYPE);
				}
				else {
					echo date('H:i:s') , " ERROR: You need to enable fileinfo or mime_magic extension.", EOL;
				}
				
				$fileContents = file_get_contents ($file);
				
				fwrite ($xmlfile,"\t\t\t<submission_file xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" stage=\"proof\" id=\"".$fileId."\" xsi:schemaLocation=\"http://pkp.sfu.ca native.xsd\">\r\n");
				
				if (!$article['fileGenre'.$i])
					$article['fileGenre'.$i] = "Article Text";
				
				fwrite ($xmlfile,"\t\t\t\t<revision number=\"1\" genre=\"".$article['fileGenre'.$i]."\" filename=\"".$article['file'.$i]."\" filesize=\"".$fileSize."\" filetype=\"".$fileType."\" uploader=\"".$uploader."\">\r\n");
				
				fwrite ($xmlfile,"\t\t\t\t<name locale=\"".$articleLocale."\">".$article['file'.$i]."</name>\r\n");				
				fwrite ($xmlfile,"\t\t\t\t<embed encoding=\"base64\">");
				fwrite ($xmlfile, base64_encode($fileContents));
				fwrite ($xmlfile,"\t\t\t\t</embed>\r\n");
				
				fwrite ($xmlfile,"\t\t\t\t</revision>\r\n");				
				fwrite ($xmlfile,"\t\t\t</submission_file>\r\n\r\n");
				# save galley data
				$galleys[$fileId] = "\t\t\t\t<name locale=\"".$fileLocale."\">".$article['fileLabel'.$i]."</name>\r\n";

				$galleys[$fileId] .= searchLocalisations('fileLabel'.$i, $article, 4, 'name');
				$galleys[$fileId] .= "\t\t\t\t<seq>".$fileSeq."</seq>\r\n";
				$galleys[$fileId] .= "\t\t\t\t<submission_file_ref id=\"".$fileId."\" revision=\"1\"/>\r\n";
				
				$fileId++;
			
			}
			if (preg_match("@^https?://@", $article['file'.$i]) && $article['file'.$i] != "") {
				# save remote galley data
				$galleys[$fileId] = "\t\t\t\t<name locale=\"".$fileLocale."\">".$article['fileLabel'.$i]."</name>\r\n";
				$galleys[$fileId] .= searchLocalisations('fileLabel'.$i, $article, 4, 'name');
				$galleys[$fileId] .= "\t\t\t\t<seq>".$fileSeq."</seq>\r\n";
				$galleys[$fileId] .= "\t\t\t\t<remote src=\"" . $article['file'.$i] . "\" />\r\n";
			}
			$fileSeq++;
		}
		
		# Submission galleys
		if (isset($galleys)){
			foreach ($galleys as $galley){
				fwrite ($xmlfile,"\t\t\t<article_galley xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" approved=\"false\" xsi:schemaLocation=\"http://pkp.sfu.ca native.xsd\">\r\n");
				fwrite ($xmlfile, $galley);
				fwrite ($xmlfile,"\t\t\t</article_galley>\r\n\r\n");				
			}
		}
	
		# pages
		if (!empty($article['pages'])){	
			fwrite ($xmlfile,"\t\t\t<pages>".$article['pages']."</pages>\r\n\r\n");
		}

		
		fwrite ($xmlfile,"\t\t</article>\r\n\r\n");
	

	}

	# After exiting the loop close the last XML file
	echo date('H:i:s') , " Closing XML file" , EOL;
	fwrite ($xmlfile,"\t\t</articles>\r\n");
	fwrite ($xmlfile,"\t</issue>\r\n\r\n");	
	fwrite ($xmlfile,"</issues>\r\n\r\n");


	echo date('H:i:s') , " Conversion finished" , EOL;


	

/* 
 * Helpers 
 * -----------
 */


# Function for searching alternative locales for a given field
function searchLocalisations($key, $input, $intend, $tag = null, $flags = null) {
    global $locales;
	
	if ($tag == "") $tag = $key;
	
	$nodes = "";
	$pattern = "/:".$key."/";
	$values = array_intersect_key($input, array_flip(preg_grep($pattern, array_keys($input), $flags)));
		
	foreach ($values as $keyval => $value){
		if ($value != ""){
			$shortLocale = explode(":", $keyval);
			if (strpos($value, "\n") !== false || strpos($value, "&") !== false || strpos($value, "<") !== false || strpos($value, ">") !== false ) $value = "<![CDATA[".nl2br($value)."]]>";
			for ($i = 0; $i < $intend; $i++) $nodes .= "\t";
			$nodes .= "<".$tag." locale=\"".$locales[$shortLocale[0]]."\">".$value."</".$tag.">\r\n";
		}
	}
	
	return $nodes;
	
}

# Function for searching alternative locales for a given taxonomy field
function searchTaxonomyLocalisations($key, $key_singular, $input, $intend, $flags = null) {
    global $locales;
		
	$nodes = "";
	$intend_string = "";
	for ($i = 0; $i < $intend; $i++) $intend_string .= "\t";
	$pattern = "/:".$key."/";
	$values = array_intersect_key($input, array_flip(preg_grep($pattern, array_keys($input), $flags)));
		
	foreach ($values as $keyval => $value){
		if ($value != ""){

			$shortLocale = explode(":", $keyval);

			$nodes .= $intend_string."<".$key." locale=\"".$locales[$shortLocale[0]]."\">\r\n";

			$subvalues = explode(";", $value);
			foreach ($subvalues as $subvalue){
				$nodes .= $intend_string."\t<".$key_singular."><![CDATA[".trim($subvalue)."]]></".$key_singular.">\r\n";	
			}

			$nodes .= $intend_string . "</".$key.">\r\n";

		}
	}
	
	return $nodes;
	
}


# Function for creating an array using the first row as keys
function createArray($sheet) {
	$highestrow = $sheet->getHighestRow();
	$highestcolumn = $sheet->getHighestColumn();
	$columncount = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestcolumn);
	$headerRow = $sheet->rangeToArray('A1:' . $highestcolumn . "1");
	$header = $headerRow[0];
	array_unshift($header,"");
	unset($header[0]);
	$array = array();
	for ($row = 2; $row <= $highestrow; $row++) {
		$a = array();
		for ($column = 1; $column <= $columncount; $column++) {
			if (strpos($header[$column], "bstract") !== false) {
					if ($sheet->getCellByColumnAndRow($column,$row)->getValue() instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText) {
						$value = $sheet->getCellByColumnAndRow($column,$row)->getValue();
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
						        }  elseif ($element->getFont()->getSuperScript()) {
						            $cellData .= '</sup>';
						        } elseif ($element->getFont()->getItalic()) {
						            $cellData .= '</em>';
						        }
						    }
						}
						$a[$header[$column]] = $cellData;
                	}
                	else{
                		$a[$header[$column]] = $sheet->getCellByColumnAndRow($column,$row)->getFormattedValue();
                	}
			}
			else {
				$key = $header[$column];
				$a[$key] = $sheet->getCellByColumnAndRow($column,$row)->getFormattedValue();
			}
		}
		$array[$row] = $a;
	}
	
	return $array;
}

# Check the highest author number
function countMaxAuthors($sheet) {
	$highestcolumn = $sheet->getHighestColumn();
	$headerRow = $sheet->rangeToArray('A1:' . $highestcolumn . "1");
	$header = $headerRow[0];
	$authorFirstnameValues = array();
	foreach ($header as $headerValue) {
		if (strpos($headerValue, "authorFirstname") !== false) {
			$authorFirstnameValues[] = (int) trim(str_replace("authorFirstname", "", $headerValue));
		}
	}
	return max($authorFirstnameValues);
}

# Check the highest file number
function countMaxFiles($sheet) {
	$highestcolumn = $sheet->getHighestColumn();
	$headerRow = $sheet->rangeToArray('A1:' . $highestcolumn . "1");
	$header = $headerRow[0];
	$fileValues = array();
	foreach ($header as $headerValue) {
		if (strpos($headerValue, "fileLabel") !== false) {
			$fileValues[] = (int) trim(str_replace("fileLabel", "", $headerValue));
		}
	}
	return max($fileValues);
}

# Function for data validation
function validateArticles($articles) {
	global $filesFolder;
	$errors = "";
	$articleRow = 1;

	foreach ($articles as $article) {

			$articleRow++;

			if (empty($article['issueYear'])) {
				$errors .= date('H:i:s') . " ERROR: Issue year missing for article " . $articleRow . EOL;
			}

			if (empty($article['issueDatepublished'])) {
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

				if (!empty($article['file'.$i]) && !preg_match("@^https?://@", $article['file'.$i]) ) {

					$fileCheck = $filesFolder.$article['file'.$i]; 

					if (!file_exists($fileCheck)) 
						$errors .= date('H:i:s') . " ERROR: file ".$i." missing " . $fileCheck . EOL;

					$fileLabelColumn = 'fileLabel'.$i;
					if (empty($fileLabelColumn)) {
						$errors .= date('H:i:s') . " ERROR: fileLabel ".$i." missing for article " . $articleRow . EOL;
					}
					$fileLocaleColumns = 'fileLocale'.$i;
					if (empty($fileLocaleColumns)) {
						$errors .= date('H:i:s') . " ERROR: fileLocale ".$i."  missingfor article " . $articleRow . EOL;
					}
				} else {
					break;
				}
			}	
	}
	
	return $errors;

}

