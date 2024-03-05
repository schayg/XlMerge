<?php
namespace ExcelMerge\Tasks;

/**
 * Modifies and copies the worksheet rels file into the target
 *
 * @package ExcelMerge\Tasks
 */
class WorksheetRels extends MergeTask { 
	public function merge($origsheetnr = null , $zip_dir=null, $file_counter=null) {
		if($origsheetnr){
			$xml_basefilename = "/xl/worksheets/_rels/";
			$target_basefilename = $this->result_dir . $xml_basefilename;
			$existing_rels = glob("{$zip_dir}/xl/worksheets/_rels/sheet{$origsheetnr}.xml.rels");
			foreach ($existing_rels as $rel) {
				$source = new \DOMDocument();
				$source->load($rel);
				$xpath = new \DOMXPath($source);
				$xpath->registerNamespace("m", "http://schemas.openxmlformats.org/package/2006/relationships");
				$elems = $xpath->query("//m:Relationship");
				foreach($elems as $e){
					$targ= $e->getAttribute('Target');
					if(strpos($targ,"xml")){$newarg=str_replace(".xml","_".$file_counter.".xml",$targ);}
					if(strpos($targ,"bin")){$newarg=str_replace(".bin","_".$file_counter.".bin",$targ);}
					$e->setAttribute('Target',$newarg);
				}
			    $target=$target_basefilename."sheet" . $this->sheet_number . ".xml.rels";
//			    $taget=$this->result_dir . "/xl/worksheets/_rels/sheet2.xml.rels";
			    $source->save($target);
			}
		}
	}
}
