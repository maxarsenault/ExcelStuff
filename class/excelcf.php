<?php

// Example of what I need to produce.
//  With variable sqref
//<conditionalFormatting sqref="I2:I4 J5"><cfRule type="colorScale" priority="10"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FF5A8AC6"/><color rgb="FFFCFCFF"/><color rgb="FFF8696B"/></colorScale></cfRule></conditionalFormatting>


class ExcelCF {

  private static $cur_priority = 20 ;

  public static function createGroup( $xml, $col, $lowrow, $highrow ) {
    
	$CFgroup = $xml->addChild( 'conditionalFormatting' ) ;
	$CFgroup->addAttribute( 'sqref', $col.$lowrow.":".$col.$highrow ) ;
	
	$CFrule = $CFgroup->addChild('cfRule') ;
	$CFrule->addAttribute( 'type', 'colorScale' ) ;
	$CFrule->addAttribute( 'priority', ExcelCF::$cur_priority ++ ) ;
	
	$colorScale = $CFrule->addChild( 'colorScale' ) ;
	
	$cfvo1   = $colorScale->addChild( 'cfvo' ) ;
	$cfvo1->addAttribute( 'type', "min" ) ;
	
	$cfvo2   = $colorScale->addChild( 'cfvo' ) ;
	$cfvo2->addAttribute( 'type', "percentile" ) ;	
	$cfvo2->addAttribute( 'val', "50" ) ;	
	
	$cfvo3   = $colorScale->addChild( 'cfvo' ) ;
	$cfvo3->addAttribute( 'type', "max" ) ;	
	
	$color1 = $colorScale->addChild( 'color' ) ;
	$color1->addAttribute( 'rgb', 'FF5A8AC6' ) ;
	
	$color2 = $colorScale->addChild( 'color' ) ;
	$color2->addAttribute( 'rgb', 'FFFCFCFF' ) ;
	
	$color3 = $colorScale->addChild( 'color' ) ;
	$color3->addAttribute( 'rgb', 'FFF8696B' ) ;
  }















}




?>