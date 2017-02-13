<?php
include "./class/excelextractor.php" ;

if ( isset($_GET['id']) && is_numeric( $_GET['id'] ) ) {
  $fetch_id = $_GET['id'] ;
} else {
  $fetch_id = 0 ;
}

if ( isset( $_FILES['fileupload']['name'] ) && trim($_FILES['fileupload']['name']) != "" ) {
  $filename = $_FILES['fileupload']['name'] ;
  $tmpfilename = $_FILES['fileupload']['tmp_name'] ;
  if ( move_uploaded_file($tmpfilename, EXCEL_UPLOAD_DIR.$filename ) ) {
    $fetch_id = ExcelExtractor::processExcel(EXCEL_UPLOAD_DIR.$filename) ;
  } else {
    echo "<div class='error'>There was an error moving " . $tmpfilename . " to " . EXCEL_UPLOAD_DIR.$filename . "</div>" ;
    die() ;
  }
}
?>
<html>
<head>
  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.1.0/jquery.min.js"></script>
  <link rel='stylesheet' type='text/css' href='./css/style.css'/>
</head>
<body>
  <input type='hidden' id="fetch_id" value="<?php echo $fetch_id ; ?>">
  <form action="<?php echo $_SERVER['PHP_SELF'] ; ?>" method="post" enctype="multipart/form-data">
    <div id="tool-bar">
      <div id="upload-div">
          <input type="file" name="fileupload" id="fileupload">
          <label class="upload-label" for="fileupload" >Upload xlsx file...</label>
      </div>
    </div>
  </form>
  <div id='table'>
  
  </div>
  <script>
    (function(){
      fetchId = $("#fetch_id").val() ;
      if ( fetchId != "" ) {
        $.getJSON( "fetchJSON.php?id=" + fetchId, function( data ){
        
          row = $("<div>").appendTo( "#table" ) ;
          abc = " ABCDEFGHIJKLMNOPQRSTUVWXYZ" ;
          a = 0; b = 0 ;
          for( i = 0 ; i <= 50 ; i++) {
            $("<div>").addClass("outside thead").text(abc[a]+abc[b]).appendTo( row ) ;
            if (b >= abc.length - 1) {
              a++ ;
              b = 1 ;
            } else {
              b++ ; 
            }
          }
        
        
         
          for( i = 0 ; i < 100 ; i++ ) {
            row = $("<div>").appendTo( "#table" ) ;
            
            $("<div>").addClass("outside tside").text(i+1).appendTo( row ) ;
            for( j = 0 ; j < 50 ; j++ ) {
              $("<div>").html("&nbsp;").appendTo( row ) ;
            }
          }
          cells = $( "#table > div > div:not(.outside)" ) ;
          $.each( data, function( key, val ){

            col = val[0] ;
            row = val[1] ;
            val = val[2] ;
            
            if ( col < 50) {
              $(cells[(row * 50) + col - 1]).text( val ) ;
            }
          }) ;
        }) ;
      }
    })();
    $("#fileupload").change(function(){
      $("form").submit();
    }) ;
    
  </script>
</body>
</html>