<?php
 ini_set('error_log', '/home/del2017/test.155del2017.it/error_log_php');
 use PhpOffice\PhpWord\Element\Section;
 use PhpOffice\PhpWord\Shared\Converter;
 use PhpOffice\PhpWord\TemplateProcessor;
 use PhpOffice\PhpWord\Shared\Html;
 use PhpOffice\PhpWord\PhpWord;
 use PhpOffice\PhpWord\SimpleType\Jc;
 use PhpOffice\PhpWord\Element\Image as ImageElement;
 use PhpOffice\PhpWord\Shared\XMLWriter;
 use PhpOffice\PhpWord\Style\Font as FontStyle;
 use PhpOffice\PhpWord\Style\Frame as FrameStyle;
 use PhpOffice\PhpWord\Writer\Word2007\Style\Font as FontStyleWriter;
 use PhpOffice\PhpWord\Writer\Word2007\Style\Image as ImageStyleWriter;
// use PhpOffice\PhpWord\IOFactory;
 use PhpOffice\PhpWord\Settings as WordSettings;
 use PhpOffice\PhpWord\Writer\Word2007\Element\Container;
 //use PhpOffice\PhpWord\Shared\Converter;
 use PhpOffice\PhpWord\Style\Font;
 
 if (isset($_POST['submit'])) {
include_once './samples/Sample_Header.php';

     $languageEnGb = new \PhpOffice\PhpWord\Style\Language(\PhpOffice\PhpWord\Style\Language::EN_GB);
 
     // Creating the new document...
     $phpWord = new \PhpOffice\PhpWord\PhpWord();
     $phpWord->getSettings()->setThemeFontLang($languageEnGb);
     /* Note: any element you append to a document must reside inside of a Section. */
 
     // Adding an empty Section to the document...
     $section = $phpWord->addSection(array('marginLeft' => 880, 'marginRight' => 880));
 
 
     //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!  this is the sector of global variables of coloring ad styling  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
     //ALERT this is just the default there are cheanges are mande on the way 
     //this is the globla font
     $fontStyletot = "'Segoe UI";
 
     //the default colors  and styles
     $defaultTextSyle = array('bold' => false,  'size' => 12, 'color' => 'dark', 'name' => $fontStyletot); // normal text with dakr color 
 
     //this is the deault default alert colorof the alert
     $defaulRED = array('bold' => false, 'size' => 12, 'color' => 'red', 'name' => $fontStyletot);
     $defaultORANGE = array('bold' => false,  'size' => 12, 'color' => 'orange', 'name' => $fontStyletot);
     $defaultGREEN = array('bold' => false,  'size' => 12, 'color' => 'green', 'name' => $fontStyletot);
 
     //bold font style
     $bold=array('bold'=>true);
 
         //this is the alerts in ceter after the chart
     $textDefaultStyle= array('bold' => false,'color'=>"black",'size'=>10,'name'=>$fontStyletot);
 
     //table style 
     $fancyTableStyleName = 'Tabelle';
     $fancyTableStyle = array('borderSize' => 1.5,  'cellMargin' => 19,'borderColor' => '4BACC6', 'alignment' => \PhpOffice\PhpWord\SimpleType\JcTable::CENTER);
     $fancyTableFirstRowStyle = array(   'borderBottomColor' => '4BACC6', 'bgColor' => '83cbde');
     $fancyTableFontStyle = array('bold' => true ,'height'=>3);
     //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!  this is the sector of global variables of coloring ad styling  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
 
 
     $section->addImage('logo.jpg', array('width' => 200, 'height' => 80, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER));
 
     //this is the header fon style 
     $phpWord->addFontStyle("headercal", array('bold' => false,  'size' => 18, 'color' => 'dark', 'name' => $fontStyletot));
     //text center code
     $phpWord->addParagraphStyle("headercal1", array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER, 'spaceAfter' => 100));
     $section->addText('"Report del Cruscotto di Controllo"', 'headercal', 'headercal1');
 
 
     //this is the second header text style
     $phpWord->addFontStyle("headersecond", $defaultTextSyle);
     $section->addText('Misurazione della permanenza in azienda della Continuità Aziendale ex Art. 2086 2` Comma Codice Civile', 'headersecond', 'headercal1');
 
     //the informative labels  regarding the datas
     $top3labels = $defaultTextSyle;
     $top3labels['size'] = 9;
     $top3labels['bold'] = true;
 
     $phpWord->addFontStyle("infohead", $top3labels);
     $section->addText('Cliente: ', 'infohead');
     $section->addText('Periodo di Riferimenti: ', 'infohead');
     $section->addText('Tipologia del Cruscotto: ', 'infohead');
 
 
     //the thirt label is on center
     $defaulRED['size'] = 35;
 
 
     $section->addText('Stato di salute generale dell`Azienda', 'headersecond', 'headercal1');
     $section->addText('31%', $defaulRED, 'headercal1');
     // Define styles
 
     //index generale graps
     $categories = array("A", "B");
     $series1 = array(80, 20);
     $style = array(
         'width'          => Converter::cmToEmu(18),
         'height'         => Converter::cmToEmu(3.6),
         '3d'             => false,
         'showAxisLabels' => false,
         'showLabels' => false,
         'showGridX'      => false,
         'showGridY'      => false,
         'colors'         => array('FF0000', 'FFFFFF', '0000FF'),
         'setCategoryLabelPosition' => 'low',
         'dataLabelOptions'=> array(
             'showVal'          => false, 
             'showCatName'      => false,  
             'showLegendKey'    => false,  
             'showSerName'      => false,  
             'showPercent'      => false,
             'showLeaderLines'  => false,
             'showBubbleSize'   => false,
         ),
     );
 
     $chart = $section->addChart('doughnut', $categories, $series1, $style);
 
 
     $paragraphStyleName = 'pStyle';
     $phpWord->addParagraphStyle($paragraphStyleName, array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER, 'spaceAfter' => 100));
     $fontStyleName = 'rStyle';
     $phpWord->addFontStyle($fontStyleName,$textDefaultStyle);
 
     $textArrayStyle1=$textDefaultStyle;
 
     $textArrayStyle1['color']="red";
     $textArrayStyle1['bold']=true;
     $section->addText('I am styled by a font style definition.', $textArrayStyle1, $paragraphStyleName);
 
     $textArrayStyle2=$textDefaultStyle;
     $textArrayStyle2['color']="green";
     $textArrayStyle2['bold']=true;
     $section->addText('I am styled by a font style definition.', $textArrayStyle2, $paragraphStyleName);
 
     //this ad 2 brake in the word file 
     $section->addTextBreak(2);
     $pagesIndexs=$defaultTextSyle;
     $pagesIndexs['size']=15;
      //this is a minitable to make the index label to look like is in row
     $table = $section->addTable();
     $table->addRow();
     $table->addCell(5000)->addText('Area finanziaria e economica', $pagesIndexs, 'headercal1');
     $table->addCell(5000)->addText('31', $pagesIndexs, 'headercal1');
 
 
  
 //remove the space into the cells of the tables
 $noSpace = array('spaceAfter' => 0,'alignment' => 'left');
 $noSpaceCentered =  $noSpace;
 $noSpaceCentered['alignment']='center';
 
 
 
 
 $phpWord->addTableStyle($fancyTableStyleName, $fancyTableStyle, $fancyTableFirstRowStyle);
 $table = $section->addTable($fancyTableStyleName);
 $table->addRow();
 $table->addCell(9900)->addText('Campo', $fancyTableFontStyle, $noSpaceCentered);
 $table->addCell(900)->addText('Valore', $fancyTableFontStyle, $noSpaceCentered);
 
 
     $table->addRow();
     $table->addCell(2000)->addText("Rispetto al periodo che stai misurando, quanti mesi sei in ritardo sulla quadratura delle banche in contabilità",array(), $noSpace);
     $table->addCell(2000)->addText("110021",array(), $noSpaceCentered);
   
     $table->addRow();
     $table->addCell(2000)->addText("Ricavi (dato riportato a 12 mesi",array(), $noSpace);
     $table->addCell(2000)->addText("13646565",array(), $noSpaceCentered);
 
     $table->addRow();
     $table->addCell(2000)->addText("Capitale investito annuo",array(), $noSpace);
     $table->addCell(2000)->addText("35456245",array(), $noSpaceCentered);
 
     $table->addRow();
     $table->addCell(2000)->addText("ROS",$bold, $noSpace);
     $table->addCell(2000)->addText("122214145",$bold, $noSpaceCentered);
 
 
     $table->addRow();
     $table->addCell(2000)->addText("TOCI",$bold, $noSpace);
     $table->addCell(2000)->addText("93",array(), $noSpaceCentered);
 
 
     $table->addRow();
     $table->addCell(2000)->addText("ROI index",$bold, $noSpace);
     $table->addCell(2000)->addText("100",array(), $noSpaceCentered);
 
  
 // Saving the document as OOXML file...
 $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
  $objWriter->save('helloWorld.docx');
 $file = 'helloWorld.docx';

 if (file_exists($file)) {
     header('Content-Description: File Transfer');
     header('Content-Type: application/octet-stream');
     header('Content-Disposition: attachment; filename="' . basename($file) . '"');
     header('Expires: 0');
     header('Cache-Control: must-revalidate');
     header('Pragma: public');
     header('Content-Length: ' . filesize($file));
     readfile($file);
     if (file_exists($file)) {
        unlink($file);
    }
     exit;
 }

 }
 
    
    
    ?>


<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.6.0/css/bootstrap.min.css"
        integrity="sha512-P5MgMn1jBN01asBgU0z60Qk4QxiXo86+wlFahKrsQf37c9cro517WzVSPPV1tDKzhku2iJ2FVgL67wG03SGnNA=="
        crossorigin="anonymous" />
    <script src="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.6.0/js/bootstrap.min.js"
        integrity="sha512-XKa9Hemdy1Ui3KSGgJdgMyYlUg1gM+QhL6cnlyTe2qzMCYm4nAZ1PsVerQzTTXzonUR+dmswHqgJPuwCq1MaAg=="
        crossorigin="anonymous"></script>
    <style>
    html,
    body,
    iframe {
        width: 100%;
        height: 100%;
        padding: 0;
        margin: 0;
    }

    </style>
</head>

<body class="jumbotron">
    <form action='' method='POST'> <button type='submit' class="btn btn-md btn-info" name='submit'>Scarica Word</button>
    </form>

</body>

</html>
