<?





class ExcelSharedStrings {

  private $shared ;
  private $sharedIndexes ;

  /*
  private static function isInitialized() {
    if ( $this->shared === "" ) {
	  die( "Run ExcelSharedStrings::openSharedStrings() first." ) ;
	}
  }
  */
  
  //public function openSharedStrings($path) {
  public function __construct($path) {
    $this->sharedIndexes = array() ;
	
    if ( ($this->shared = simplexml_load_file($path)) === false ) {
	  die( "Failed to open up shared strings xml, path: " . $path ) ;
	}
	
	$this->initSharedStringIndexes() ;
  }

  public function getCellValue( $cell ) {

    //ExcelSharedStrings::isInitialized() ;
  
	if ( !is_object($cell) ) {
	  error_log( "Returning void." ) ;
	  return ;
	}
  
    if ( trim($cell->attributes()->t) == "s" ) {
	  $cell_value = trim($cell->v[0]) + 0 ;
	  return trim($this->shared->si[ $cell_value ]->t[0]) ;
	} else {
	  return trim($cell->v[0]) ;
	}
  
  }



  private function initSharedStringIndexes() {
  
    //ExcelSharedStrings::isInitialized() ;
  
    foreach ( $this->shared->si as $stringIndex => $si ) {

	  $string = trim($si->t[0]) ;

	  if ( $string === "" ) {
	    continue ;
	  }

	  // error_log( $string ) ;
	  $this->sharedIndexes[ $string ] = $stringIndex ;
	}
  }
  
  
  
  public function sharedStringIndex($string) {
    $string = trim( $string ) ;
    if ( isset( $this->sharedIndexes[$string] ) ) {
	  return $this->sharedIndexes[$string] ;
	} else {
	  return false ;
	} 
  }

  public function getSharedString( $index ) {
    return $this->shared->si[$index]->t[0] ;
  }
  
  // Check if string is in xml,
  //  add string to xml and return index of the new one.
  public function addSharedString( $string ) {
    // error_log( $string . " = string to add" ) ;  
  
    $si = $this->shared->addChild('si') ;
	
	$si->addChild( 't', trim($string) ) ;
	$indexnum = count( $this->shared ) - 1 ;
	$this->sharedIndexes[$string] = $indexnum ;
	return $indexnum ;
  }
  
  public function writeToFile( $path ) {
    $handle = fopen($path,  "w");
    if (fwrite( $handle, $this->shared->asXML() ) ) {
	  
	} else {
	  die( "FAILED TO WRITE TO: " . $path ) ;
	}
    fclose( $handle ) ;
  }
}

?>