<?


//include_once $_SERVER['DOCUMENT_ROOT'] . "/exceltool/Classes/PHPExcel.php" ;

class ExcelSegmentRow extends ExcelChartRow {

  private $rawData ; 
  private $type ;
  private $colStart ;
  private $colEnd ;
  
  private $cells ; 
  private $title ;
  
  function __construct( $row, $sharedStringsHandler ) {

    $this->rawData = $row ;
	$this->title   = "" ;

	$rowId  = trim($row->attributes()->r) ;	
	
	// Set $this->type 
	// required

	
	if ( !isset($row->c[1]) ) {
	//  die( "Missing tag on line: " . $rowId ) ;
	}
    $cell = $row->c[1] ;
  
  
    //error_log( "B col" ) ;
    if ( ExcelFunctions::getCellCol( $rowId, $cell ) == "B" ) {
	  $this->type = strtolower($sharedStringsHandler->getCellValue( $cell )) ;
	}
	
	
	// Set $this->range
	// not required
	
	
    //error_log( "C col" ) ;
	if ( isset($row->c[3]) ) {
	
	  $cell = $row->c[3] ;
      
      if ( ExcelFunctions::getCellCol( $rowId, $cell ) == "D" ) {
	  
	    $range = explode( ':', $sharedStringsHandler->getCellValue($cell) ) ;
		if ( count( $range ) == 1 ) {
		
		  // error_log( "G: " . 'D'  ) ;
		  $this->colStart = PHPExcel_Cell::columnIndexFromString( 'E' ) - 1 ;
		  if ( trim($range[0]) !== "" ) {
		    // error_log( "H: " . trim( $range[0] )  ) ;
		    $this->colEnd   = PHPExcel_Cell::columnIndexFromString( trim($range[0]) ) ;
		  } else {
		    // error_log( "I: " . "ZZ" ) ;
		    $this->colEnd   = PHPExcel_Cell::columnIndexFromString( "ZZ" ) ;
		  }
		} else if ( count( $range ) == 2 ) {
		  // error_log( "J: " . trim($range[0])  ) ;
		  $this->colStart = PHPExcel_Cell::columnIndexFromString( trim($range[0]) ) - 1 ;
		  // error_log( "K: " . trim($range[1])  ) ;
		  $this->colEnd   = PHPExcel_Cell::columnIndexFromString( trim($range[1]) ) ;
		}
	  }
	}
	
	
	$this->cells = array() ;
	
	foreach( $row->c as $key => $cell ) {
	  $col = ExcelFunctions::getCellCol( $rowId, $cell ) ;
	  // error_log( "K: " . $col  ) ;
	  $col = PHPExcel_Cell::columnIndexFromString( $col ) ;
	  if ( $col > 2 ) {
	    $this->cells[ $col ] = $cell ;
		//error_log( "CELL VALUE: " . $cell->v[0] ) ;
		
		//error_log("col".$this->cells[$col]->attributes()->r);
	  }
	 // error_log( $col. ":" . ExcelSharedStrings::getCellValue($cell) ) ;
	}
  }
  
  
  public function getTitleCell() {
  
    if ( $this->type == 'title' ) {
	
	  foreach( $this->cells as $cell ) {
	  
	  	if ( $cellKey > $this->colStart && $cellKey <= $this->colEnd ) {
		  // 
	    } else if ( trim($this->colEnd) !== "" ) {
	      continue ;
	    }
		
	    if ( ExcelFunctions::getCellValue( $cell ) !== "" ) {
		  return $cell ;
		}
	  }
	}
    return "" ;
  }
  
  
  public function appendToXML( $sheet, $sharedDest, $sharedSource, $currentRowIndex  ) {
	
    $row = $sheet->addChild('row') ;

    $row->addAttribute('r', $currentRowIndex) ;
    $row->addAttribute('x14ac:dyDescent', '0.25') ;
    $row->addAttribute('spans', '1:22') ;

	
	foreach( $this->cells as $cellKey => $cell ) {
	
	  if ( $cellKey > $this->colStart && $cellKey <= $this->colEnd ) {
		// 
	  } else if ( trim($this->colEnd) !== "" ) {
	    continue ;
	  }
	  
	  $newcell = $row->addChild('c') ;
		
	  $newcell->addChild('v',$cell->v[0]) ;
		
	  foreach ( $cell->attributes() as $key => $value ) {
		if ( trim($key) == "t" && trim($value) == "s" ) {
		  if ( ($stringIndex = $sharedDest->sharedStringIndex( $stringValue = $sharedSource->getCellValue( $cell ) ))  !== false ) {
			$newcell->v[0] = $stringIndex ;
			//error_log( "String found outcome: " . $stringValue . ", index = " . $stringIndex  ) ;
		  } else {
			$newcell->v[0] = $sharedDest->addSharedString( $stringValue ) ;
			//error_log( "String found outcome: " . $stringValue . ", index = " . $newcell->v[0]  ) ;
		  }		    
		}
		$newcell->addAttribute( $key, $value ) ;
	  }
		
		
	  $colLetter = ExcelFunctions::stripCellCol( $cell ) ;
	  $colLetter = PHPExcel_Cell::stringFromColumnIndex( PHPExcel_Cell::columnIndexFromString($colLetter) - 4 ) ;
	  $newcell['r'] = $colLetter . $currentRowIndex ;
		
	  
	}

  
  }
}





?>