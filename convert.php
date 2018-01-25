<?php

/* 
 * Settings 
 * ---------
 */

// The file containing the metadata
$fileName = 'scripta-1970.xlsx';

// The default locale. For other locales use locale:fieldName
$defaultLocale = 'en_US';

// The uploader account name
$uploader = "admin";

// The maximum number of authors per article, eg. authorLastname3 => 3
$maxAuthors = 1;

// The maximum number of files per article, eg. file2 => 2
$maxFiles = 1;

// Set to '1' if you only want to validate the data
$onlyValidate = 0;

// Location of full text files
$filesFolder = dirname(__FILE__) . "/files/";

// Possible locales
$locales = array(
				'en' => 'en_US',
				'fi' => 'fi_FI',
				'sv' => 'sv_SE',
				'de' => 'de_DE',
				'ru' => 'ru_RU',
				'fr' => 'fr_FR',
			);

// PHPExcel settings
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/Helsinki');
define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');
require_once dirname(__FILE__) . '/phpexcel/Classes/PHPExcel.php';


/* 
 * Load Excel data to an array
 * ------------------------------------
 */
echo date('H:i:s') , " Creating a new PHPExcel object" , EOL;
$objReader = PHPExcel_IOFactory::createReaderForFile($fileName);
$objReader->setReadDataOnly(); 
$objPHPExcel = $objReader->load($fileName);

echo date('H:i:s') , " Creating an array" , EOL;
$sheet = $objPHPExcel->setActiveSheetIndex(0);
$articles = createArray($sheet);


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
foreach ($articles as $key => $article){
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
		
		# Issue identification
		fwrite ($xmlfile,"\t\t<issue_identification>\r\n");
		
		if ($article['issueVolume'])
			fwrite ($xmlfile,"\t\t\t<volume>".$article['issueVolume']."</volume>\r\n");	
		if ($article['issueNumber'])
			fwrite ($xmlfile,"\t\t\t<number>".$article['issueNumber']."</number>\r\n");			
		fwrite ($xmlfile,"\t\t\t<year>".$article['issueYear']."</year>\r\n");
		
		if ($article['issueTitle']){
			fwrite ($xmlfile,"\t\t\t<title>".$article['issueTitle']."</title>\r\n");
		}
		# Add alternative localisations for the issue title
		fwrite ($xmlfile, searchLocalisations('issueTitle', $article, 3));
		
		fwrite ($xmlfile,"\t\t</issue_identification>\r\n\r\n");
		
		fwrite ($xmlfile,"\t\t<date_published>".$article['issueDatepublished']."</date_published>\r\n\r\n");
		
		# Sections
		fwrite ($xmlfile,"\t\t<sections>\r\n");
		    
			foreach ($sections[$article['issueDatepublished']] as $sectionAbbrev => $sectionTitle){
				fwrite ($xmlfile,"\t\t\t<section ref=\"".$sectionAbbrev."\">\r\n");
				fwrite ($xmlfile,"\t\t\t\t<abbrev locale=\"".$defaultLocale."\">".$sectionAbbrev."</abbrev>\r\n");
				fwrite ($xmlfile,"\t\t\t\t<title locale=\"".$defaultLocale."\">".$sectionTitle."</title>\r\n");
				fwrite ($xmlfile, searchLocalisations('issueTitle', $article, 3));
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

	fwrite ($xmlfile,"\t\t<article xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" locale=\"".$defaultLocale."\" stage=\"production\" date_published=\"".$article['issueDatepublished']."\" section_ref=\"".$article['sectionAbbrev']."\">\r\n\r\n");
	
		# Title, subtitle, Abstract
		fwrite ($xmlfile,"\t\t\t<title locale=\"".$defaultLocale."\">".$article['title']."</title>\r\n");
		fwrite ($xmlfile, searchLocalisations('title', $article, 3));
		
		if (isset($article['subtitle'])){	
			fwrite ($xmlfile,"\t\t\t<subtitle locale=\"".$defaultLocale."\">".$article['subtitle']."</subtitle>\r\n");
		}
		fwrite ($xmlfile, searchLocalisations('subtitle', $article, 3));
		
		if (isset($article['abstract'])){
			fwrite ($xmlfile,"\t\t\t<abstract locale=\"".$defaultLocale."\"><![CDATA[".$article['abstract']."]]></abstract>\r\n\r\n");
		}
		fwrite ($xmlfile, searchLocalisations('abstract', $article, 3));
	
		# TODO: add support for licence, supporting agencies and keywords
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
			
			if ($article['authorLastname'.$i]){
				
				if ($i == 1)
					fwrite ($xmlfile,"\t\t\t\t<author primary_contact=\"true\" include_in_browse=\"true\" user_group_ref=\"Author\">\r\n");
				else
					fwrite ($xmlfile,"\t\t\t\t<author include_in_browse=\"true\" user_group_ref=\"Author\">\r\n");
				
				fwrite ($xmlfile,"\t\t\t\t\t<firstname>".$article['authorFirstname'.$i]."</firstname>\r\n");
				fwrite ($xmlfile,"\t\t\t\t\t<lastname>".$article['authorLastname'.$i]."</lastname>\r\n");

				if (isset($article['authorAffiliation'.$i])){
					fwrite ($xmlfile,"\t\t\t\t\t<affiliation locale=\"".$defaultLocale."\">".$article['authorAffiliation'.$i]."</affiliation>\r\n");
				}
				fwrite ($xmlfile, searchLocalisations('authorAffiliation'.$i, $article, 5, 'affiliation'));

				fwrite ($xmlfile,"\t\t\t\t\t<email><![CDATA[]]></email>\r\n");
				
				if (isset($article['country'.$i])){
					fwrite ($xmlfile,"\t\t\t\t\t<country>".$article['country'.$i]."</country>\r\n");
				}
				if (isset($article['orcid'.$i])){
					fwrite ($xmlfile,"\t\t\t\t\t<orcid>".$article['orcid'.$i]."</orcid>\r\n");
				}
				if (isset($article['authorBio'.$i])){
					fwrite ($xmlfile,"\t\t\t\t\t<biography locale=\"".$defaultLocale."\"><![CDATA[".$article['authorBio'.$i]."]]></biography>\r\n");
				}
				
				fwrite ($xmlfile,"\t\t\t\t</author>\r\n");

				
			}

		}

		fwrite ($xmlfile,"\t\t\t</authors>\r\n\r\n");
	
	
		# Submission files
		unset($galleys);
		$fileSeq = 0;
		
		for ($i = 1; $i <= $maxFiles; $i++) {
			
			if ($article['file'.$i]){
					
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
				
				fwrite ($xmlfile,"\t\t\t\t<revision number=\"1\" genre=\"".$article['fileGenre'.$i]."\" filename=\"".$article['file'.$i]."\" filesize=\"".$fileSize."\" filetype=\"".$fileType."\" user_group_ref=\"Author\" uploader=\"".$uploader."\">\r\n");
				
				fwrite ($xmlfile,"\t\t\t\t<name locale=\"".$defaultLocale."\">".$article['file'.$i]."</name>\r\n");				

				fwrite ($xmlfile,"\t\t\t\t<embed encoding=\"base64\">");
				fwrite ($xmlfile, base64_encode($fileContents));
				fwrite ($xmlfile,"\t\t\t\t</embed>\r\n");
				
				fwrite ($xmlfile,"\t\t\t\t</revision>\r\n");				
				fwrite ($xmlfile,"\t\t\t</submission_file>\r\n\r\n");

				# save galley data
				$galleys[$fileId] = "\t\t\t\t<name locale=\"".$defaultLocale."\">".$article['fileLabel'.$i]."</name>\r\n";
				$galleys[$fileId] .= searchLocalisations('fileLabel'.$i, $article, 4, 'name');
				$galleys[$fileId] .= "\t\t\t\t<seq>".$fileSeq."</seq>\r\n";
				$galleys[$fileId] .= "\t\t\t\t<submission_file_ref id=\"".$fileId."\" revision=\"1\"/>\r\n";
				
				$fileId++;
				$fileSeq++;
				
			}

		}
	
	
		# Submission galleys
		foreach ($galleys as $key => $galley){
			fwrite ($xmlfile,"\t\t\t<article_galley xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" approved=\"false\" xsi:schemaLocation=\"http://pkp.sfu.ca native.xsd\">\r\n");
			fwrite ($xmlfile, $galley);
			fwrite ($xmlfile,"\t\t\t</article_galley>\r\n\r\n");				
		}
	
		# pages
		if (isset($article['pages'])){	
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
			if (strpos($value, "\n") !== false) $value = "<![CDATA[".nl2br($value)."]]>";
			for ($i = 0; $i < $intend; $i++) $nodes .= "\t";
			$nodes .= "<".$tag." locale=\"".$locales[$shortLocale[0]]."\">".$value."</".$tag.">\r\n";
		}
	}
	
	return $nodes;
	
}

# Function for creating an array using the first row as keys
function createArray($sheet) {
	$highestrow = $sheet->getHighestRow();
	$highestcolumn = $sheet->getHighestColumn();
	$columncount = PHPExcel_Cell::columnIndexFromString($highestcolumn);
	$header = $sheet->rangeToArray('A1:' . $highestcolumn . "1");
	$body = $sheet->rangeToArray('A2:' . $highestcolumn . $highestrow);
	
	$array = array();
	for ($row = 0; $row <= $highestrow - 2; $row++) {
		$a = array();
		for ($column = 0; $column <= $columncount - 1; $column++) {
			$a[$header[0][$column]] = $body[$row][$column];
		}
		$array[$row] = $a;
	}
	
	return $array;
	
}

# Function for searching empty values
function emptyElementExists($arr) {
	return array_search("", $arr) !== false;
}

# Function for sorting the data by pubdate
function sortByIssueDate($a, $b) {
	return strcasecmp($a['issueDatepublished'], $b['issueDatepublished']);
}

# Function for data validation
function validateArticles($articles) {
	
	global $maxFiles, $filesFolder, $defaultLocale;
	$errors = "";
	
	if (emptyElementExists(array_column($articles, 'issueYear'))){
		$errors .= date('H:i:s') . " ERROR: Issue year missing" . EOL;
	}

	if (emptyElementExists(array_column($articles, 'issueDatepublished'))){
		$errors .= date('H:i:s') . " ERROR: Issue publication date missing" . EOL;
	}

	if (emptyElementExists(array_column($articles, 'title'))){
		$errors .= date('H:i:s') . " ERROR: article title missing for the given default locale ". $defaultLocale . EOL;
	}

	if (emptyElementExists(array_column($articles, 'sectionTitle'))){
		$errors .= date('H:i:s') . " ERROR: section title missing for the given default locale " . $defaultLocale . EOL;
	}

	if (emptyElementExists(array_column($articles, 'sectionAbbrev'))){
		$errors .= date('H:i:s') . " ERROR: section abbreviation missing for the given default locale " . $defaultLocale . EOL;
	}

	foreach ($articles as $key => $article){
			for ($i = 1; $i < $maxFiles; $i++) {
				if ($article['file'.$i]){
					
					$fileCheck = $filesFolder.$article['file'.$i]; 
					
					if (!file_exists($fileCheck)) 
						$errors .= date('H:i:s') . " ERROR: file missing " . $fileCheck . EOL;
				}
			}	
	}
	
	return $errors;

}


