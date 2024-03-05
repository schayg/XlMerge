<?php
namespace ExcelMerge\Tasks;

/**
 * Consolidates the contents of two 'xl/calcChains.xml' files into one
 * Schayg
 * @package ExcelMerge\Tasks
 */
class CalcChains extends MergeTask { 
public function merge($zip_dir = null) {
		$xml_filename = "/xl/calcChain.xml";
		$target_filename = $this->result_dir . $xml_filename;
		$source_filename = $zip_dir . $xml_filename;
		if(!file_exists($target_filename)){ // there is no calcChains yet, then just copy over the one from the source if there is one.
			if(file_exists($source_filename)){copy($source_filename,$target_filename);}
		}
		if(file_exists($source_filename) && file_exists($target_filename)){
			$target = new \DOMDocument();
			$target->load($target_filename);
			$source = new \DOMDocument();
			$source->load($source_filename);
			$xpatht = new \DOMXPath($target);
			$xpatht->registerNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			$elems = $xpatht->query("//m:c");
			$imax=1;
			foreach ($elems as $e) {
				$i = $e->getAttribute("i");
			if(intval($i)>$imax){$imax=intval($i);}
			}
			$xpath = new \DOMXPath($source);
			$xpath->registerNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			$elems = $xpath->query("//m:c");
			foreach ($elems as $e) {
				$i = $e->getAttribute("i");
				$e->setAttribute("i",intval($i)+$imax);
			$ne=$target->importNode($e,true);
			$target->documentElement->appendChild($ne);
			}
			$target->save($target_filename);
			return $imax;
		} else {return null;}
	}
}
