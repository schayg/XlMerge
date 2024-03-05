<?php
namespace XlMerge;

/**
 * Merges two or more Excel files into one.
 *
 * Only Excel 2007 files are supported, so you can only merge .xlsx and .xlxm
 * files. So far, it only seems to work with files that are generated with
 * PHPExcel.
 *
 * @author schayg https://github.com/schayg
 * based heavilt on infostreams/excel-merge
 * @package XlMerge
 * @property $working_dir
 * @property $result_dir
 */
class ExcelMerge {
	protected $files       = array();
	private   $working_dir = null;
	private   $tmp_dir     = null;
	private   $result_dir  = null;
	private   $tasks;
	private   $file_counter = null;  // schayg: we count the files, this is needed to avoid naming conflict in the drawings,charts and relations to them.
	public    $debug       = false;
	protected $names       = array(); // schayg this holds the sheet naming basenames, every file will have sheets with consecutive numering + basename in the merged file.
	private   $bsn         = null;
	public function __construct($files = array(), $names = array()) {
		// create a temporary directory with an understandable name
		// (comes in use when debugging)
		// -> SCHG: uniqid IS already a timestamp, lets strip the repetition, and make names shorter.
		for ($i=0; $i < 5; $i++) {
			$this->working_dir =
				sys_get_temp_dir() .
				DIRECTORY_SEPARATOR . "xlsx_merge". DIRECTORY_SEPARATOR .
				uniqid() .
				DIRECTORY_SEPARATOR;

			if (!is_dir($this->working_dir)) {
				mkdir($this->working_dir, 0755, true);
				break;
			}
		}

		if (!is_dir($this->working_dir)) {
			trigger_error("Could not create temporary working directory {$this->working_dir}", E_USER_ERROR);
		}


		$this->tmp_dir = $this->working_dir . "tmp" . DIRECTORY_SEPARATOR;
		mkdir($this->tmp_dir, 0755, true);

		$this->result_dir = $this->working_dir . "result" . DIRECTORY_SEPARATOR;
		mkdir($this->result_dir, 0755, true);

		$this->registerMergeTasks();
		$this->file_counter=1;
		foreach ($files as $f) {
			$i=$this->file_counter;
			$this->bsn = $names[$i-1];
			$this->addFile($f);
			$this->addObjects($f);
			$this->file_counter++;
		}
	}

	public function __destruct() {
		if (!$this->debug) {
			$this->removeTree(realpath($this->working_dir));
		}
	}

	protected function addObjects($zip_dir=null){ // extension by Schayg to handle graphs, drawings
		//we will simply copy over the relevant xml files, and to avoid file name collision, we will include the file counter in their names.
		// have to register them afterwards in the content types file.
		// So we also have to change the names in the "rels" files.
		//rels locations:  
			// 		/xl/worksheets/_rels  -> sheetname.xml.rels ** this is handled as a separate task, NOT here, as we need to rename the rel according to the new sheet names.
			//		/xl/charts/_rels  -> chartname.xml.rels -> refers to styleX.xml and colorsX.xml and possibly to ../drawings/...
			//		/xl/drawings/_rels -> drawingname.xml.rels -> refers to charts or drawings
		if(!is_dir("{$this->result_dir}/xl/charts")){mkdir("{$this->result_dir}/xl/charts",0755,true);}
		if(!is_dir("{$this->result_dir}/xl/charts/_rels")){mkdir("{$this->result_dir}/xl/charts/_rels",0755,true);}
		if(!is_dir("{$this->result_dir}/xl/drawings")){mkdir("{$this->result_dir}/xl/charts",0755,true);}
		if(!is_dir("{$this->result_dir}/xl/drawings/_rels")){mkdir("{$this->result_dir}/xl/charts/_rels",0755,true);}
		if(!is_dir("{$this->result_dir}/xl/printerSettings")){mkdir("{$this->result_dir}/xl/printerSettings",0755,true);}
		$existing_xmls=glob("{$zip_dir}/xl/drawings/drawing*.xml");
		foreach($existing_xmls as $x){
			$last = basename($x);
			sscanf($last, "drawing%d.xml", $number);
			$n="/xl/drawings/drawing".$number."_".$this->file_counter.".xml";
			copy($x,$this->result_dir.$n);
			$ctf = "{$this->result_dir}/[Content_Types].xml";
			$dom = new \DOMDocument();
			$dom->load($ctf);
			$tag = $dom->createElement("Override");
			$tag->setAttribute('PartName',$n);
			$tag->setAttribute('ContentType', "application/vnd.openxmlformats-officedocument.drawing+xml");
			$dom->documentElement->appendChild($tag);
			$dom->save($ctf);
		}
		$existing_xmls=glob("{$zip_dir}/xl/drawings/_rels/*.rels");
		foreach($existing_xmls as $x){
			$last = basename($x);
			sscanf($last, "drawing%d.xml.rels", $number);
			$targname=$this->result_dir."/xl/drawings/_rels/drawing".$number."_".$this->file_counter.".xml.rels";
			$source = new \DOMDocument();
			$source->load($x);
			$xpath = new \DOMXPath($source);
			$xpath->registerNamespace("m", "http://schemas.openxmlformats.org/package/2006/relationships");
			$elems = $xpath->query("//m:Relationship");
			foreach($elems as $e){
				$targ= $e->getAttribute('Target');
				if(strpos($targ,"xml")){$newarg=str_replace('.xml','_'.$this->file_counter.'.xml',$targ);}
				$e->setAttribute('Target',$newarg);
				}
			$source->save($targname);
			}
		$existing_xmls=glob("{$zip_dir}/xl/charts/chart*.xml");
		foreach($existing_xmls as $x){
			$last = basename($x);
			sscanf($last, "chart%d.xml", $number);
			$n="/xl/charts/chart".$number."_".$this->file_counter.".xml";
			copy($x,$this->result_dir.$n);
			$ctf = "{$this->result_dir}/[Content_Types].xml";
			$dom = new \DOMDocument();
			$dom->load($ctf);
			$tag = $dom->createElement("Override");
			$tag->setAttribute('PartName',$n);
			$tag->setAttribute('ContentType', "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
			$dom->documentElement->appendChild($tag);
			$dom->save($ctf);
		}
		$existing_xmls=glob("{$zip_dir}/xl/charts/colors*.xml");
		foreach($existing_xmls as $x){
			$last = basename($x);
			sscanf($last, "colors%d.xml", $number);
			$n="/xl/charts/colors".$number."_".$this->file_counter.".xml";
			copy($x,$this->result_dir.$n);
			$ctf = "{$this->result_dir}/[Content_Types].xml";
			$dom = new \DOMDocument();
			$dom->load($ctf);
			$tag = $dom->createElement("Override");
			$tag->setAttribute('PartName',$n);
			$tag->setAttribute('ContentType', "application/vnd.ms-office.chartcolorstyle+xml");
			$dom->documentElement->appendChild($tag);
			$dom->save($ctf);
		}
		$existing_xmls=glob("{$zip_dir}/xl/charts/style*.xml");
		foreach($existing_xmls as $x){
			$last = basename($x);
			sscanf($last, "style%d.xml", $number);
			$n="/xl/charts/style".$number."_".$this->file_counter.".xml";
			copy($x,$this->result_dir.$n);
			$ctf = "{$this->result_dir}/[Content_Types].xml";
			$dom = new \DOMDocument();
			$dom->load($ctf);
			$tag = $dom->createElement("Override");
			$tag->setAttribute('PartName',$n);
			$tag->setAttribute('ContentType', "application/vnd.ms-office.chartstyle+xml");
			$dom->documentElement->appendChild($tag);
			$dom->save($ctf);
		}
		$existing_xmls=glob("{$zip_dir}/xl/charts/_rels/*.rels");
		foreach($existing_xmls as $x){
			$last = basename($x);
			sscanf($last, "chart%d.xml.rels", $number);
			$targname=$this->result_dir."/xl/charts/_rels/chart".$number."_".$this->file_counter.".xml.rels";
			$source = new \DOMDocument();
			$source->load($x);
			$xpath = new \DOMXPath($source);
			$xpath->registerNamespace("m", "http://schemas.openxmlformats.org/package/2006/relationships");
			$elems = $xpath->query("//m:Relationship");
			foreach($elems as $e){
				$targ= $e->getAttribute("Target");
				if(strpos($targ,"xml")){$newarg=str_replace(".xml","_".$this->file_counter.".xml",$targ);}
				$e->setAttribute("Target",$newarg);
				}
			$source->save($targname);
			}
		$existing_bins=glob("{$zip_dir}/xl/printerSettings/*.bin");
		foreach($existing_bins as $x){
			$last = basename($x);
			sscanf($last, "Settings%d.bin", $number);
			$last=str_replace(".bin","_".$this->file_counter.".bin",$last);
			copy($x,"{$this->result_dir}/xl/printerSettings/{$last}");
		}
	}
	public function addFile($filename) {
		if ($this->isSupportedFile($filename)) {
			if ($this->resultsDirEmpty()) {
				$this->addFirstFile($filename);
			} else {
				$this->mergeWorksheets($filename);
			}
//			$this->files[] = $filename; // schayg: add back the filename?? why? tha array is already parsed by the constructor, this here is not needed any more, skipped.
		}
	}


	/**
	 * Saves the merged file.
	 *
	 * @param null $where
	 * @return string The path and filename to the saved file. The file extension can be
	 * different from the one you provided (!)
	 */
	public function save($where = null) {
		$zipfile = $this->zipContents();
		if ($where === NULL) {
			$where = $zipfile;
		}

		// ignore whatever extension the user might have given us and use the one
		// we obtained in 'zipContents' (i.e. either XLSX or XLSM)
		$where =
			pathinfo($where, PATHINFO_DIRNAME) .
			DIRECTORY_SEPARATOR .
			pathinfo($where, PATHINFO_FILENAME) . "." .
			pathinfo($zipfile, PATHINFO_EXTENSION);

		// move the zipped file to the provided destination
		rename($zipfile, $where);

		// returns the name of the file
		return $where;
	}

	/**
	 * Downloads the merged file
	 *
	 * @param null $download_filename
	 */
	public function download($download_filename = null) {
		$zipfile = $this->zipContents();
		if ($download_filename === NULL) {
			$download_filename = $zipfile;
		}

		// ignore whatever extension the user might have given us and use the one
		// we obtained in 'zipContents' (i.e. either XLSX or XLSM)
		$download_filename =
			pathinfo($download_filename, PATHINFO_FILENAME) . "." .
			pathinfo($zipfile, PATHINFO_EXTENSION);

		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment;filename="' . $download_filename . '"');
		header('Cache-Control: max-age=0');
		echo file_get_contents($zipfile);
		unlink($zipfile);
		die;
	}

	protected function scanCharts($path=null,$oldname=null,$newname=null){
	    if(is_dir($path) && oldaname && newname){
		// $chartspath=$this->result_dir."/xl/charts/";
		$charts=glob("{$path}/chart*");
		foreach($charts as $ch){
		    $chstr=file_get_contents($ch);
		    if($chstr=str_replace($oldname, $newname,$chstr)) file_put_contents($ch, $chstr); // if there were replacements then re-write charts file.
		    }
	    }
	}
		protected function scanSheets($path=null,$oldname=null,$newname=null){
	    if(is_dir($path) && oldaname && newname){
		$charts=glob("{$path}/sheet*");
		foreach($charts as $ch){
		    $chstr=file_get_contents($ch);
		    if($chstr=str_replace($oldname, $newname,$chstr)) file_put_contents($ch, $chstr); // if there were replacements then re-write charts file.
		    }
	    }
	}

	protected function addFirstFile($filename) {
		if ($this->resultsDirEmpty()) {
			if ($this->isSupportedFile($filename)) {
				$this->unzip($filename, $this->result_dir);
			// schayg: we have to change the worksheet names here too in the workbook.xml
			$wbfilename = "{$this->result_dir}/xl/workbook.xml";
			$dom = new \DOMDocument();
			$dom->load($wbfilename);
			$xpath = new \DOMXPath($dom);
			$xpath->registerNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			$elems = $xpath->query("//m:sheet");
			foreach ($elems as $e) {
			    $osn=$e->getAttribute('name');
			    $sn=$this->bsn.'_'.$osn;
			    $e->setAttribute('name', $sn);
			    $this->scanCharts($this->result_dir."/xl/charts",$osn,$sn); // scan charts xml-s and re-write sheet references to new name. Excel stupidly does not use RefId for sheet reference, or sheet xml name, but the name which is in the workbook.xml.
				$this->scanSheets($this->result_dir."/xl/worksheets",$osn,$sn); // scan worksheets, as they also can contain references with Sheetnames....
			    }
			$dom->save($wbfilename);
			if(file_exists($this->result_dir."/xl/calcChain.xml")){unlink($this->result_dir."/xl/calcChain.xml");}
			// done: schayg
			}
		} else {
			$this->mergeWorksheets($filename);
		}
	}


	protected function mergeWorksheets($filename) {
		if ($this->resultsDirEmpty()) {
			$this->addFirstFile($filename);
		} else {
			if ($this->isSupportedFile($filename)) {
				$zip_dir = $this->tmp_dir . DIRECTORY_SEPARATOR . basename($filename);
				$this->unzip($filename, $zip_dir);
				// remap chart references to new names
				$shared_strings = $this->tasks->sharedStrings->merge($zip_dir);
				list($styles, $conditional_styles) = $this->tasks->styles->merge($zip_dir);
				$this->tasks->vba->merge($zip_dir);
				$wbfilename = "{$zip_dir}/xl/workbook.xml";
				$dom = new \DOMDocument();
				$dom->load($wbfilename);
				$xpath = new \DOMXPath($dom);
				$xpath->registerNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
				$elems = $xpath->query("//m:sheet");
				foreach ($elems as $e) {
				    $osn=$e->getAttribute('name');
				    $sn=$this->bsn.'_'.$osn;
				    $this->scanCharts($zip_dir."/xl/charts",$osn,$sn); // scan charts xml-s and re-write sheet references to new name.
					$this->scanSheets($zip_dir."/xl/worksheets",$osn,$sn); // scan worksheets and re-write references to new worksheet name.
				    }
				// $this->tasks->CalcChains->merge($zip_dir); // schayg extension
				$this->addObjects($zip_dir);
				$worksheets = glob("{$zip_dir}/xl/worksheets/sheet*.xml");
				foreach ($worksheets as $s) {
					list($sheet_number, $sheet_name, $is_hidden) = $this->tasks->worksheet->merge($s, $shared_strings, $styles, $conditional_styles);
					$origsheetnr=null;
					$last = basename($s);
					sscanf($last, "sheet%d.xml", $origsheetnr);
					if ($sheet_number!==false) {
						$this->tasks->workbookRels->set($sheet_number, $sheet_name,$is_hidden)->merge();
						$this->tasks->contentTypes->set($sheet_number, $sheet_name,$is_hidden)->merge();
						$this->tasks->app->set($sheet_number, $sheet_name,$is_hidden)->merge();
						// $nc=$this->file_counter-1;$bn=($this->names)[$nc];
						$this->tasks->workbook->set($sheet_number, $sheet_name,$is_hidden)->merge($this->bsn);
						$this->tasks->worksheetRels->set($sheet_number,$sheet_name,$is_hidden)->merge($origsheetnr,$zip_dir,$this->file_counter);
//						$this->scanCharts($zip_dir."/xl/charts",$sheet_name,"{$this->bsn}_{$sheet_name}"); // remap chart references
					}
				}
			}
		}
	}

	protected function registerMergeTasks() {
		$this->tasks = new \stdClass();

		// global tasks
		$this->tasks->sharedStrings = new Tasks\SharedStrings($this);
		$this->tasks->styles = new Tasks\Styles($this);
		$this->tasks->vba = new Tasks\Vba($this);

		// worksheet tasks
		$this->tasks->worksheet = new Tasks\Worksheet($this);
		$this->tasks->workbookRels = new Tasks\WorkbookRels($this);
		$this->tasks->contentTypes = new Tasks\ContentTypes($this);
		$this->tasks->app = new Tasks\App($this);
		$this->tasks->workbook = new Tasks\Workbook($this);
		$this->tasks->worksheetRels = new Tasks\WorksheetRels($this); // schayg extension
//		$this->tasks->CalcChains = new Tasks\CalcChains($this); // schayg extension
	}


	protected function isSupportedFile($filename, $throw_error = true) {
		$ext = pathinfo($filename, PATHINFO_EXTENSION);
		$is_supported = in_array(strtolower($ext), array('xlsx', 'xlsm'));
		if (!$is_supported && $throw_error) {
			user_error("Can only merge Excel files in .XLSX or .XLSM format. Skipping " . $filename, E_USER_WARNING);
		}

		return $is_supported;
	}

	protected function resultsDirEmpty() {
		return count(array_diff(scandir($this->result_dir), array('.', '..'))) == 0;
	}


	protected function unzip($filename, $directory) {
		$zip = new \ZipArchive();
		$zip->open($filename);
		$zip->extractTo($directory);
		$zip->close();
	}

	protected function removeTree($dir) {
		$result = false;

		$dir = realpath($dir);
		if (strpos($dir, realpath(sys_get_temp_dir())) === 0) {
			$result = true;
			$files = array_diff(scandir($dir), array('.', '..'));
			foreach ($files as $file) {
				if (is_dir("$dir/$file")) {
					$result &= $this->removeTree("$dir/$file");
				} else {
					$result &= unlink("$dir/$file");
				}
			}
			$result &= rmdir($dir);
		}

		return $result;
	}

	protected function zipContents() {
		$zip_directory = realpath($this->result_dir);
		$target_zip = $this->working_dir . DIRECTORY_SEPARATOR . "merged-excel-file";
		$ext = "xlsx";

		$delete = array();

		$zip = new \ZipArchive();
		$zip->open($target_zip, \ZipArchive::CREATE | \ZipArchive::OVERWRITE);

		// Create recursive directory iterator
		/** @var \SplFileInfo[] $files */
		$files = new \RecursiveIteratorIterator(
			new \RecursiveDirectoryIterator($zip_directory),
			\RecursiveIteratorIterator::LEAVES_ONLY
		);

		foreach ($files as $name => $file) {
			// Skip directories (they would be added automatically)
			if (!$file->isDir()) {
				// Get real and relative path for current file
				$filePath = $file->getRealPath();
				if (basename($filePath) != $target_zip) {
					$relativePath = substr($filePath, strlen($zip_directory) + 1);

					// Add current file to archive
					$zip->addFile($filePath, $relativePath);

					$delete[] = $filePath;

					if (basename($filePath) == "vbaProject.bin") {
						// we found VBA code; we change the extension to 'XLSM' to enable macros
						$ext = "xlsm";
					}
				}
			}
		}

		// Zip archive will be created only after closing object
		$zip->close();

		// by default, we delete the files that we put in the zip file
		if (!$this->debug) {
			foreach ($delete as $d) {
				unlink($d);
			}
		}

		// give the zipfile its final name
		rename($target_zip, "$target_zip.$ext");

		return "$target_zip.$ext";
	}

	public function __get($name) {
		switch ($name) {
			case "result_dir":
				return $this->result_dir;
			case "working_dir":
				return $this->working_dir;
		}
		return null;
	}
}