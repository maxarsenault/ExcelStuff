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

  private static $dbh = -1 ;

  private static function dbConnect() {
    if ( ExcelExtractor::$dbh == -1 ) {
      try {
        ExcelExtractor::$dbh = new PDO('mysql:host=maxdb.c6zdviphb64q.us-west-2.rds.amazonaws.com:3306;dbname=Portfolio', "QAZWSX", "f1Lx!n438RNp");
      } catch (PDOException $e){
        print "Error!: " . $e->getMessage() . "<br/>" ;
        die() ;
      }
    }
    return ExcelExtractor::$dbh ;
  }
  
  public static function storeExcel( $json ){
    $conn = ExcelExtractor::dbConnect() ;
    
    $stmt = $conn->prepare("INSERT INTO excel_json (excel_id, json_string) values(DEFAULT, :value) ;") ;
    $stmt->bindParam( ":value", $json ) ;
    $stmt->execute();
    
    return $conn->lastInsertId() ;
  }
  
  public static function retrieveExcel( $id ){
    $conn = ExcelExtractor::dbConnect() ;
    
    $stmt = $conn->prepare("SELECT json_string FROM excel_json WHERE excel_id=:id ;") ;
    $stmt->bindParam( ":id", $id ) ;
    $stmt->execute();
    
    return $stmt->fetch(PDO::FETCH_ASSOC)["json_string"];
  }
  
  // For packing an xlsx file.
  /*   For use in future projects
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
  */

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


 public static function processExcel( $filename ) {
    if ( $filename != "" ) {
      ExcelExtractor::extractZip( $filename, REF_PATH ) ; 
      
      $sheetname = "sheet1" ;
      
      $sheetfile = ExcelFunctions::findSheet( REF_PATH."xl/workbook.xml", $sheetname ) ;
      
      // Load the xlsx to be parsed
      $xmlB = simplexml_load_file(REF_PATH."xl/worksheets/".$sheetfile);
      $sharedB = new ExcelSharedStrings(REF_PATH."xl/sharedStrings.xml") ;
      
      $chart = ExcelFunctions::collectChart($xmlB, $sharedB) ;
      $json = $chart->chartAsJSON() ;
      return ExcelExtractor::storeExcel( $json );
    }
  }
}
?>