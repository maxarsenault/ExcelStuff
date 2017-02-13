<?php

include_once $_SERVER['DOCUMENT_ROOT'] . "/exceltool8/class/excelchartrow.php";

class ExcelChart {

  private static $lists = array() ;

  private $rows ;
  private $sharedStringHandler ;
  private $labelsIndex ;
  private $title ;
  
  function __construct( $shared ) {

    $this->rows = array() ;
    $this->sharedStringHandler = $shared ;
    $this->labelsIndex = 0 ;
    $this->title = "" ;
  }

  public static function addList( $row, $shared ) {
  
    $listType = $shared->getCellValue( $row->c[1] ) ;
	
	if ( !isset( $lists[$listType] ) ) {
	  $lists[$listType] = array() ;
	}
	
    for ( $i = 2 ; $i < count( $row->c ) ; $i++ ) {
      $value = $shared->getCellValue( $row->c[$i] ) ;
      foreach( explode(',',$value) as $val ) {
          ExcelChart::$lists[ $listType ] = trim($val) ;
      }
    }
  }
  
  public static function getLists( $listType ) {
    return ExcelChart::$lists[ $listType ] ;
  }
  
  public function chartAsJSON() {
    $jsonRows = array() ;
  
    foreach( $this->rows as $row ) {
      foreach( $row->getCells() as $key => $str ) {
        
        $jsonRows[] = array( $key, $row->getRowNum(), $str ) ;
      }
    }
    
    return json_encode( $jsonRows ) ;
  }
  
  public function addRow( $row ) {
    $this->rows[] = new ExcelChartRow( $row, $this->sharedStringHandler ) ;
  }

  
  public function appendToXML( $xml, $sharedDest ) {
 
    $sheet = $xml->sheetData[0] ;
	
	
	$currentRowIndex = trim( $sheet->row[ count($sheet->row) - 1 ]->attributes()->r ) + 2 ;
  
    foreach( $this->rows as $row ) {
	
      $row->appendToXML( $sheet, $sharedDest, $this->sharedStringHandler, $currentRowIndex ) ;
	  $currentRowIndex++ ;
	}
  
  }
  
  private function findLabelIndexes( $row, $list, $shared ) {
  
    $colIndexList = array() ;
  
  	$colIndexList['E'] = 1 ;
  
    $key = 1 ;
	
	$row_length = count($row->c) ;
    for ( $i = $row_length - 1 ; $i >= 0 ; $i-- ) {
	  if ( ($cellVal = $shared->getCellValue( $row->c[$i] )) != "" && isset( $list[ $cellVal ] ) ) {
	    $colIndexList[ ExcelFunctions::stripCellCol( $row->c[$i] ) ] = 1 ;
	  }
	}
	
	return $colIndexList ;
  
  }
  
  
  public function appendStringToXML( $string, $color, $sheet, $sharedDest, $sharedSource, $currentRowIndex, $validColList ) {
  
    $row = $sheet->addChild('row') ;
    $row->addAttribute('r', $currentRowIndex) ;
    $row->addAttribute('x14ac:dyDescent', '0.25') ;
    $row->addAttribute('spans', '1:22') ;
	
    $destCol = 1 ;
	
    for ( $i = 0 ; $i < count($validColList) ; $i++ ) {
	
      $cell = $row->addChild('c', "") ;
      $cell->addAttribute( 'r', PHPExcel_Cell::stringFromColumnIndex( $destCol ) . $currentRowIndex ) ;
      if ( $color !== "" ) {
        $cell->addAttribute( 's', $color ) ;
      }
      
      if ( $i == 0 ) {
        $index = $sharedDest->sharedStringIndex( $string) ;
        if ( $index === false ) {
          $index = $sharedDest->addSharedString( $string ) ;
        }
  
        $cell->addChild( 'v', $index ) ;
        $cell->addAttribute( 't', 's' ) ;
      }
      
      $destCol++ ;
    }
  
  } 
}





?>