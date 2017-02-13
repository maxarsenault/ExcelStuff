<?php
include "./class/excelextractor.php" ;

header('Content-type: application/json');
if ( isset( $_GET['id'] ) && is_numeric( $_GET['id'] ) ) {
  echo ExcelExtractor::retrieveExcel( $_GET['id'] ) ;
} else {
  echo encode_JSON( array() ) ;
}

?>