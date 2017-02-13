<?



// I need to make this not a static library.

class ExcelSharedStrings {

  private static $shared = "" ;
  private static $sharedIndexes = array() ;

  
  private static function isInitialized() {
    if ( ExcelSharedStrings::$shared === "" ) {
	  die( "Run ExcelSharedStrings::openSharedStrings() first." ) ;
	}
  }
  
  public static function openSharedStrings($path) {
    ExcelSharedStrings::$shared = simplexml_load_file($path);  
  }

  public static function getCellValue( $cell ) {

    ExcelSharedStrings::isInitialized() ;
  
    if ( trim($cell->attributes()->t) == "s" ) {
	  $cell_value = trim($cell->v[0]) + 0 ;
	  return trim(ExcelSharedStrings::$shared->si[ $cell_value ]->t[0]) ;
	} else {
	  return trim($cell->v[0]) ;
	}
  
  }



  public static function initSharedStringIndexes() {
  
    ExcelSharedStrings::isInitialized() ;
  
    foreach ( ExcelSharedStrings::$shared->si as $stringIndex => $string ) {
	  ExcelSharedStrings::$sharedIndexes[ $string ] = $stringIndex ;
	}
  }
  
  
  public static function isSharedString($string) {
    if ( isset( ExcelSharedStrings::$sharedIndexes[$string] ) ) {
	  return true ;
	} else {
	  return false ;
	}
  }

  public static function getSharedString( $index ) {
    return ExcelSharedStrings::$shared->si[$index]->t[0] ;
  }
  
  // Check if string is in xml,
  //   if it is, return index.
  //   if not, add string to xml and return index of the new one.
  public static function addSharedString( $string ) {
    $si = ExcelSharedStrings::$shared->addChild('si') ;
	$si->addChild( 't', $string ) ;
	return count( ExcelSharedStrings::$shared ) - 1 ;
  }
}

?>