<?

include_once $_SERVER['DOCUMENT_ROOT'] . "/exceltool5/class/excelchartrow.php";

class ExcelWorkBook {

  static public function findSheet( $file, $search ) {

    $xml = simplexml_load_file($workingpath."xl/workbook.xml");
    $search = implode( "", explode( " ", urldecode( $search )) ) ;
  
    $count = 1 ;
    foreach ($xml->sheets[0]->sheet as $sheet ) {
  
      $name = $sheet->attributes()->name ;
      $name = implode( "", explode( " ", $name ) ) ;
  
      if ( $name == $search ) {
	     ;
	  }
	  $count++ 
    }
	
	return false ;
  }
}





?>