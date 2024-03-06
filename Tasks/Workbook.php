<?php
namespace XlMerge\Tasks;

/**
 * Modifies the "xl/workbook.xml" file to contain one more worksheet.
 *
 * @package XlMerge\Tasks
 */
class Workbook extends MergeTask {
	public function merge($basename=null) {
		/**
		 * 	7. xl/workbook.xml
		 *         => add
		 *            <sheet name="{New sheet}" sheetId="{N}" r:id="rId{N}"/>
		 */
		if(!$basename){$basename=$this->sheet_number;}
		$filename = "{$this->result_dir}/xl/workbook.xml";
		$dom = new \DOMDocument();
		$dom->load($filename);

		$xpath = new \DOMXPath($dom);
		$xpath->registerNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		$elems = $xpath->query("//m:sheets");
		foreach ($elems as $e) {
			$tag = $dom->createElement('sheet');
// **Schayg: ************** the reason for "native excel' fail is that the Sheet NAMES also must be unique! That is not checked in the code.
// to have a workaround, we append the sheet number to the sheet name.
			$tag->setAttribute('name', $basename.'_'.($this->sheet_name));
// end of fix by SchayG
			$tag->setAttribute('sheetId', $this->sheet_number);
			$tag->setAttribute('r:id', "rId" . $this->sheet_number);
			if($this->is_hidden){$tag->setAttribute('state',"hidden");}  // SCHG: handle hidden sheets here.

			$e->appendChild($tag);
			break;
		}

		// make sure all worksheets have the correct rId - we might have assigned them new ids
		// in the Tasks\WorkbookRels::merge() method
		
		// Caroline Clep: this is breaking the result file - need to make sure we don't touch the sheets ids and only update the external links
		//$elems = $xpath->query("//m:sheets/m:sheet");
		//foreach ($elems as $e) {
		//	$e->setAttribute("r:id", "rId" . ($e->getAttribute("sheetId")));
		//}
		
		$relfilename = "{$this->result_dir}/xl/_rels/workbook.xml.rels";
		$reldom = new \DOMDocument();
		$reldom->load($relfilename);

		$relxpath = new \DOMXPath($reldom);
		$relxpath->registerNamespace("m", "http://schemas.openxmlformats.org/package/2006/relationships");
		$relelems = $relxpath->query("//m:Relationship");


		$elems = $xpath->query("//m:externalReference");
		$refId = 1;
		foreach ($elems as $e)
		{
			foreach ($relelems as $rele)
			{
				if ($rele->getAttribute("Target") === "externalLinks/externalLink" . $refId . ".xml")
				{
					$e->setAttribute("r:id", $rele->getAttribute("Id"));
					break;
				}
			}
			$refId++;
		}
		// Caroline Clep: End of fix

		$dom->save($filename);
	}

}
