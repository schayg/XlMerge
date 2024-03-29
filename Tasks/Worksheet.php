<?php
namespace XlMerge\Tasks;

/**
 * Adds a new worksheet to the merged Excel file
 *
 * @package XlMerge\Tasks
 */
class Worksheet extends MergeTask {

	/**
	 * Adds a new worksheet to the merged Excel file
	 *
	 * @param null $filename The filename of the sheet to copy
	 * @param array $shared_strings_mapping
	 * @param array $styles_mapping
	 * @return array
	 */
	public function merge($filename=null, $shared_strings_mapping=array(), $styles_mapping=array(), $conditional_styles_mapping=array()) {
		if (!file_exists($filename)) {
			return array(false, false);
		}
		$new_sheet_number = $this->getSheetCount($this->result_dir) + 1;

		// copy file into place
		$new_name = $this->result_dir . "/xl/worksheets/sheet{$new_sheet_number}.xml";
		if (!is_dir(dirname($new_name))) {
			mkdir(dirname($new_name));
		}
		copy($filename, $new_name);

		// adjust references to any shared strings
		$sheet = new \DOMDocument();
		$sheet->load($new_name);

		$this->remapSharedStrings($sheet, $shared_strings_mapping);
		$this->remapStyles($sheet, $styles_mapping);
		$this->remapConditionalStyles($sheet, $conditional_styles_mapping);
		$this->remapColsStyles($sheet, $styles_mapping); 
		// schayg: clear selection of the tab
		$xpath = new \DOMXPath($sheet);
		$xpath->registerNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		$elements=$xpath->query("//m:sheetViews/m:sheetView");
		foreach($elements as $e) {$e->setAttribute("tabSelected","0");}
		// save worksheet with adjustments
		$sheet->save($new_name);

		// extract worksheet name
		[$sheet_name,$is_hidden] = $this->extractWorksheetName($filename);

		return [$new_sheet_number, $sheet_name,$is_hidden];
	}

	protected function getSheetCount($dir) {
		$existing_sheets = glob("{$dir}/xl/worksheets/sheet*.xml");

		if (count($existing_sheets)>0) {
			natsort($existing_sheets);
			$last = basename(end($existing_sheets));

			if (sscanf($last, "sheet%d.xml", $number)) {
				return $number;
			}
		}

		return 0;
	}

	protected function remapSharedStrings($sheet, $mapping) {
		$xpath = new \DOMXPath($sheet);
		$xpath->registerNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		$shared = $xpath->query("//m:c[@t='s']/m:v");

		if (!is_null($shared)) {
			foreach ($shared as $tag) {
				$old_id = $tag->nodeValue;

				if (is_numeric($old_id)) {
					$old_id = intval($old_id);
					if (array_key_exists($old_id, $mapping)) {
						$tag->nodeValue = $mapping[$old_id];
					}
				}
			}
		}
	}

	protected function remapStyles($sheet, $mapping) {
		$this->doRemapping($sheet, "//m:c[@s]", "s", $mapping);
	}

	protected function remapConditionalStyles($sheet, $mapping) {
		$this->doRemapping($sheet, "//m:conditionalFormatting/m:cfRule[@dxfId]", "dxfId", $mapping);
	}

	protected function remapColsStyles($sheet, $mapping) {
		$this->doRemapping($sheet, "//m:col[@Style]", "style", $mapping);
	}
	
	protected function doRemapping($sheet, $xpath_query, $attribute, $mapping) {
		// adjust references to styles
		$xpath = new \DOMXPath($sheet);
		$xpath->registerNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		$conditional_styles = $xpath->query($xpath_query);

		if (!is_null($conditional_styles)) {
			foreach ($conditional_styles as $tag) {
				$old_id = $tag->getAttribute($attribute);

				if (is_numeric($old_id)) {
					$old_id = intval($old_id);
					if (array_key_exists($old_id, $mapping)) {
						$tag->setAttribute($attribute, $mapping[$old_id]);
					}
				}
			}
		}
	}

	protected function extractWorksheetName($filename) {
		$sheet_name=null; $is_hidden=null;
/*		$workbook = new \DOMDocument();
		$workbook->load(dirname($filename) . "/../workbook.xml");

		$xpath = new \DOMXPath($workbook);
		$xpath->registerNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		sscanf(basename($filename), "sheet%d.xml", $number);

//		$sheet_name = "Worksheet $number";
		$elems = $xpath->query("//m:sheets/m:sheet[@sheetId='" . $number . "']");
//		$elems = $xpath->query("//m:sheets/m:sheet[@sheetId='" . $sheet_number . "']");
		foreach ($elems as $e) {
			// should be one only
			$sheet_name = $e->getAttribute('name');
			break;
		}
*/ 
// mod SCHG: "native" excel may NOT have sheetId = sheetX.xml name! we MUST search and link from workbook.xml.rels through Relationship ID-s...
		$workbookrel = new \DOMDocument();
		$workbookrel->load(dirname($filename) . "/../_rels/workbook.xml.rels");
		$xpath = new \DOMXPath($workbookrel);
		$xpath->registerNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");
		$oursheets=$xpath->query("//r:Relationship[@Target='worksheets/".basename($filename)."']");
		foreach($oursheets as $sh) {
			$relid=$sh->getAttribute('Id');
		}
		$workbook = new \DOMDocument();
		$workbook->load(dirname($filename) . "/../workbook.xml");
		$xpath = new \DOMXPath($workbook);
		$xpath->registerNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
//		$xpath->registerNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");
		$oursheets=$xpath->query("//m:sheets/m:sheet[@r:id='".$relid."']");
		foreach($oursheets as $sh) {
			$sheet_name=$sh->getAttribute('name');
			$is_hidden=$sh->getAttribute('state');
		}
		return array($sheet_name,$is_hidden);
	}
}
