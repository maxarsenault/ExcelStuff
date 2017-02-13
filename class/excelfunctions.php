<?php


class ExcelFunctions {

  static private $current_title = "" ;


  
  static public function getSheet( $xml ) {
    return $xml->sheetData[0] ;
  }
  
  static public function findRowByNumber( $sheet, $search ) {
  
    $rows = $sheet->children() ;
    $rowslength = count( $rows ) ;  
  
    $search = trim( $search ) ;
  
    for ( $i = 0 ; $i < $rowslength ; $i++ ) {
	  if ( isset($rows[$i]['r']) && $rows[$i]['r'] == $search ) {
	    return $rows[$i] ;
	  }
	}
    return false ;
  }

  static public function findSheet( $file, $search ) {

    $xml = simplexml_load_file($file);
    $search = implode( "", explode( " ", urldecode( $search )) ) ;
  
    $count = 1 ;
    foreach ($xml->sheets[0]->sheet as $sheet ) {
  
      $name = $sheet->attributes()->name ;
      $name = implode( "", explode( " ", $name ) ) ;
  
      //error_log( strtolower($name) . ":" . strtolower($search) ) ;
  
      if ( strtolower($name) == strtolower($search) ) {
	    //error_log( "success" ) ;
	    return "sheet". $count . ".xml" ;
	  }
      $count++ ;
    }

    // Sheet not found
    return false ;
  }

  // $id is the full excel id, letter + number.
  //   Example: "D24"
  public static function clearCellById( $row, $id ) {
  
    foreach($row->c as $cell) {
  
      if($cell['r'] == $id) {
        $dom=dom_import_simplexml($cell);
        $dom->parentNode->removeChild($dom);
      }
    }
  }
  
  // $id is just the row number.
  //   Example: "24"
  
  // Currently corrupts the file in excel
  public static function clearRowById( $sheetData, $id ) {
  
    foreach($sheetData->row as $row) {
  
      if($row['r'] == $id) {
        $dom=dom_import_simplexml($row);
        $dom->parentNode->removeChild($dom);
      }
    }
  }
  
  public static function getCellCol( $rowId, $cell ) {
    $cellId = ExcelFunctions::getCellId( $cell ) ;
    return trim(str_replace( $rowId, '', $cellId )) ;
  }
  
  // Gets the non numeric part of a cell id
  public static function stripCellCol( $cell ) {
   
    $str = "" ;
	$colId = $cell['r'] ;
	
	for ( $i = 0 ; $i < strlen( $colId ) ; $i++ ) {
	
	  $substring = substr( $colId, $i, 1 ) ;
	
	  if ( !is_numeric( $substring ) ) {
	    $str .= $substring ;
	  } else{
	    return $str ;
	  }
	}
	return $str ;
  }
  
  public static function getCellId( $cell ) {
    if ( $cell != null ) {
      return trim($cell->attributes()->r) ;
    } else {
      return -1 ;
    }
  }
  
  public static function getRowId( $row ) {
    return trim($row->attributes()->r) ;
  }
  
  public static function addCell() {
  
  }

  public static function collectChart( $xml, $sharedStrings ) {

    $chart = new ExcelChart( $sharedStrings ) ;
	
    foreach( $xml->sheetData[0]->children() as $rowkey => $row ) {
	
      if ( count($row->children()) == 0 ) {
        continue ;
      } else {
        $chart->addRow( $row ) ;
      }
    }
    
    return $chart ;
  }

  public static function getCurrentTitle() {
    return ExcelFunctions::$current_title ;
  }

  public static function setCurrentTitle( $title ) {
    ExcelFunctions::$current_title = $title ;
  }
  
  public static function addRowToXML( $sheet, $currentRowIndex ) {

    $row = $sheet->addChild('row') ;
    $row->addAttribute('r', $currentRowIndex) ;
    $row->addAttribute('x14ac:dyDescent', '0.25') ;
    $row->addAttribute('spans', '1:22') ;
	
	return $row ;
  }
  
  public static function setCellValue( $cell, $value, $sharedDest ) {
    if ( is_numeric( $value ) ) {
      $cell->v[0] = $value ;
  	} else if ( ($stringIndex = $sharedDest->sharedStringIndex( $value )) !== false ) {
      $cell->v[0] = $stringIndex ;	
    } else {
      $cell->v[0] = $sharedDest->addSharedString( $value ) ;
    }
  }
  
  
  public static function addCellToRow( $row, $sharedDest, $x_index, $value, $style = "", $formula = "" ) {
  
    $value = trim($value) ;
  	$newcell = $row->addChild('c') ;
	
    $colLetter = PHPExcel_Cell::stringFromColumnIndex( $x_index - 1 ) ;
    $currentRowIndex = $row->attributes()->r ;
    $newcell->addAttribute( 'r', $colLetter . $currentRowIndex ) ;
	
    //error_log( $colLetter . $currentRowIndex ) ;
	
    if ( !is_numeric( $value ) ) {
    
      if ( ($stringIndex = $sharedDest->sharedStringIndex( $value )) !== false ) {
        $newcell->addChild('v', $stringIndex ) ;
      
      //error_log( "String found outcome: " . $value . ", index = " . $stringIndex  ) ;
      } else {
      $newcell->addChild('v', $stringIndex = $sharedDest->addSharedString( $value ) ) ;
      //error_log( "String added outcome: " . $value . ", index = " . $stringIndex  ) ;
      }    
    
        $newcell->addAttribute( 't', 's' ) ;
      if ( $style !== "" ) {
          $newcell->addAttribute( 's', $style ) ;
      }
    } else {
      $newcell->addChild('v', $value ) ;
      if ( $style !== "" ) {
          $newcell->addAttribute( 's', $style ) ;
      }
    }
 
    return $newcell ;
  }
  
  // Borrowed from PHPExcel
	public static function columnIndexFromString($pString = 'A') {
		static $_indexCache = array();

		if (isset($_indexCache[$pString]))
			return $_indexCache[$pString];

		static $_columnLookup = array(
			'A' => 1, 'B' => 2, 'C' => 3, 'D' => 4, 'E' => 5, 'F' => 6, 'G' => 7, 'H' => 8, 'I' => 9, 'J' => 10, 'K' => 11, 'L' => 12, 'M' => 13,
			'N' => 14, 'O' => 15, 'P' => 16, 'Q' => 17, 'R' => 18, 'S' => 19, 'T' => 20, 'U' => 21, 'V' => 22, 'W' => 23, 'X' => 24, 'Y' => 25, 'Z' => 26,
			'a' => 1, 'b' => 2, 'c' => 3, 'd' => 4, 'e' => 5, 'f' => 6, 'g' => 7, 'h' => 8, 'i' => 9, 'j' => 10, 'k' => 11, 'l' => 12, 'm' => 13,
			'n' => 14, 'o' => 15, 'p' => 16, 'q' => 17, 'r' => 18, 's' => 19, 't' => 20, 'u' => 21, 'v' => 22, 'w' => 23, 'x' => 24, 'y' => 25, 'z' => 26
		);

		if (isset($pString{0})) {
			if (!isset($pString{1})) {
				$_indexCache[$pString] = $_columnLookup[$pString];
				return $_indexCache[$pString];
			} elseif(!isset($pString{2})) {
				$_indexCache[$pString] = $_columnLookup[$pString{0}] * 26 + $_columnLookup[$pString{1}];
				return $_indexCache[$pString];
			} elseif(!isset($pString{3})) {
				$_indexCache[$pString] = $_columnLookup[$pString{0}] * 676 + $_columnLookup[$pString{1}] * 26 + $_columnLookup[$pString{2}];
				return $_indexCache[$pString];
			}
		}
		throw new PHPExcel_Exception("Column string index can not be " . ((isset($pString{0})) ? "longer than 3 characters" : "empty"));
	}

  // Borrowed from PHPExcel  
	public static function stringFromColumnIndex($pColumnIndex = 0) {
		static $_indexCache = array();

		if (!isset($_indexCache[$pColumnIndex])) {
			// Determine column string
			if ($pColumnIndex < 26) {
				$_indexCache[$pColumnIndex] = chr(65 + $pColumnIndex);
			} elseif ($pColumnIndex < 702) {
				$_indexCache[$pColumnIndex] = chr(64 + ($pColumnIndex / 26)) .
											  chr(65 + $pColumnIndex % 26);
			} else {
				$_indexCache[$pColumnIndex] = chr(64 + (($pColumnIndex - 26) / 676)) .
											  chr(65 + ((($pColumnIndex - 26) % 676) / 26)) .
											  chr(65 + $pColumnIndex % 26);
			}
		}
		return $_indexCache[$pColumnIndex];
	}  
  

}


?>