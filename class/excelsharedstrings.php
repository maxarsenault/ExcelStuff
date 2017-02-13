<?php



 

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
	  //return trim($this->shared->si[ md5($cell_value) ]->t[0]) ;
	  return trim($this->shared->si[ $cell_value ]->t[0]) ;
	} else {
	  return trim($cell->v[0]) ;
	}
  
  }



  private function initSharedStringIndexes() {
  
    //ExcelSharedStrings::isInitialized() ;
  
    //$stringCount = 0 ;
  
    foreach ( $this->shared->si as $stringIndex => $si ) {
	
	  $string = trim($si->t[0]) ;

	  if ( $string === "" ) {
	    continue ;
	  }
	  
	  //echo $stringIndex . " --- " . $si . "<br/>" ;

	  // error_log( $string ) ;
	  //$this->sharedIndexes[ md5($string) ] = $stringIndex ;
	  $this->sharedIndexes[ $string ] = $stringIndex ;
	  //$stringCount++ ;
	}
  }
  
  
  
  public function sharedStringIndex($string) {
    $string = trim( $string ) ;
    // if ( isset( $this->sharedIndexes[md5($string)] ) ) {
    if ( isset( $this->sharedIndexes[$string] ) ) {
	  // return $this->sharedIndexes[md5($string)] ;
	  return $this->sharedIndexes[$string] ;
	} else {
	  return false ;
	} 
  }

  public function getSharedString( $index ) {
    // return $this->shared->si[md5($index)]->t[0] ;
    return $this->shared->si[$index]->t[0] ;
  }
  
  // Check if string is in xml,
  //  add string to xml and return index of the new one.
  public function addSharedString( $string ) {
    // error_log( $string . " = string to add" ) ;  
  

  
    if (( $index = $this->sharedStringIndex( $string )) !== false ) {
	  return $index ;
	}
  
    $si = $this->shared->addChild('si') ;
	
	// Need to escape this some how
	$si->addChild( 't', trim($string) ) ;
	
	$indexnum = count( $this->shared ) - 1 ;
	//$this->sharedIndexes[md5(trim($string))] = $indexnum ;
	$this->sharedIndexes[trim($string)] = $indexnum ;
    // echo "ADDING STRING: " . $string . "... index is: " . $indexnum . " <br/>" ;
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