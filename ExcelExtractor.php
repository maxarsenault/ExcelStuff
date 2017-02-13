<?php


define( 'REF_PATH', '/var/www/extract/' ) ;
define( 'EXCEL_UPLOAD_DIR', '/var/www/upload/' ) ;

include_once $_SERVER['DOCUMENT_ROOT'] . "/exceltool8/Classes/PHPExcel.php" ;
include_once $_SERVER['DOCUMENT_ROOT'] . "/exceltool8/class/excelchart.php" ;
include_once $_SERVER['DOCUMENT_ROOT'] . "/exceltool8/class/excelsharedstrings.php" ;
include_once $_SERVER['DOCUMENT_ROOT'] . "/exceltool8/class/excelfunctions.php" ;
include_once $_SERVER['DOCUMENT_ROOT'] . "/exceltool8/class/excelcf.php" ;


//exec( "rm -r ".WRITE_PATH."*" ) ;
exec( "rm -r ".REF_PATH."*" ) ;


class ExcelExtractor {

  // For packing an xlsx file.
  private static function addDirectoryToZip($zip, $dir, $base)
  {
      $newFolder = str_replace($base, '', $dir);
      $zip->addEmptyDir($newFolder);
      foreach(glob($dir . '/*') as $file)
      {
          if(is_dir($file))
          {
              $zip = addDirectoryToZip($zip, $file, $base);
          }
          else
          {
              $newFile = str_replace($base, '', $file);
              $zip->addFile($file, $newFile);
        
          }
      }
      return $zip;
  }

  // for extracting an xlsx
  private static function extractZip( $openpath, $extractpath ) {
    $zip = new ZipArchive;
    $res = $zip->open( $openpath );
    if ($res === TRUE) {
      $zip->extractTo( $extractpath );
      $zip->close(); 
    } else {
      echo 'failed, code:' . $res;
      die() ;
    }
  }  


 public static function readExcel( $filename ) {
    if ( isset( $filename ) ) {
      extractZip( EXCEL_UPLOAD_DIR . $filename, REF_PATH ) ; 
      
      if ( !isset($_POST['sheetname']) || trim($_POST['sheetname']) == "" ) {
        $sheetname = "sheet1" ;
      } else {
        $sheetname = $_POST['sheetname'] ;
      }
      
      $sheetfile = ExcelFunctions::findSheet( REF_PATH."xl/workbook.xml", $sheetname ) ;
      
      if ( $sheetfile === false ) {
        die( 'Invalid sheet name: "' . $sheetname . '"' ) ; 
      }
      
      // Load the xlsx to be parsed
      $xmlB = simplexml_load_file(REF_PATH."xl/worksheets/".$sheetfile);
      $sharedB = new ExcelSharedStrings(REF_PATH."xl/sharedStrings.xml") ;
      
    } else if ( isset( $getsetnumber ) && $getsetnumber > 0 && $_POST['out'] != 5 && $_POST['out'] != 7 && $_POST['out'] != 9 && $_POST['out'] != 8 ) {

      error_log( "Writing DB to file B" ) ;
      extractZip( 'templatev8.xlsx', REF_PATH ) ;
      
      //$xmlB = simplexml_load_file(REF_PATH."xl/worksheets/".$sheetfile);
      $xmlB = simplexml_load_file(REF_PATH."xl/worksheets/sheet1.xml");
      $sharedB = new ExcelSharedStrings(REF_PATH."xl/sharedStrings.xml") ;

      ExcelFunctions::appendToXMLFromDB( $getsetnumber, $xmlB, $sharedB ) ;

    }
  }


  // Segment charts and none
  error_log( "COLLECTING START" ) ;
  $chart = ExcelFunctions::collectChart($xmlB, $sharedB) ;
  header('Content-Type: application/json'); 
  echo $chart->chartAsJSON() ;


  //$chart->appendSegmentChart( $xml, $sharedA ) ;






  if ( !isset( $_POST['out'] ) || $_POST['out'] == 0 ) {
    /*
    unset( $_POST ) ;
    include "index.php" ;*/
    die() ;
  }

  /*
  $handle = fopen(WRITE_PATH."xl/worksheets/sheet1.xml", "w");
  fwrite( $handle, $xml->asXML() ) ; 
  fclose( $handle ) ;

  if ( $_POST['out'] == 2 && isset( $getsetnumber ) ) {
   
    $sharedB->writeToFile(WRITE_PATH."xl/sharedStrings.xml") ;
  } else {
    $sharedA->writeToFile(WRITE_PATH."xl/sharedStrings.xml") ;
  }


  $newfile = WRITE_PATH."newexcel.xlsx" ;
  $rels = WRITE_PATH . "_rels" ;
  $doc  = WRITE_PATH . "docProps" ;
  $xl   = WRITE_PATH . "xl" ;
  $rootfile = WRITE_PATH . "[Content_Types].xml" ;
  $relsfile = WRITE_PATH . "_rels/.rels" ;

  $zip = new ZipArchive ;
  $res = $zip->open( $newfile, ZipArchive::CREATE ) ;

  addDirectoryToZip($zip, $rels, WRITE_PATH ) ;
  addDirectoryToZip($zip, $doc, WRITE_PATH ) ;
  addDirectoryToZip($zip, $xl, WRITE_PATH ) ;
  $zip->addFile( $rootfile, "[Content_Types].xml" ) ;
  $zip->addFile( $relsfile, "_rels/.rels" ) ;

  $zip->close(); 

  header('Content-Disposition: attachment;filename=exceltest.xlsx');
  header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  readfile( $newfile ) ;

  exec( " rm " . $newfile, $output ) ;
  */


}
?>