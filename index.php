<?php

session_start();
require_once('vendor/autoload.php');

function estilo($word){

  $phpWord = $word;
  $phpWord->setDefaultFontName('Arial');
  $phpWord->setDefaultFontSize(11);

  $phpWord->addTitleStyle(1, array(
    'size'=>16, 
    'color'=>'8B0000'
  ));

  $phpWord->addFontStyle('titulo', array(
    'color' => '8B0000',
    'bold' => true,
    'size' => 16,
  ));

  $phpWord->addParagraphStyle('titulo1', array(
    'align' => 'center',
  ));

  $phpWord->addParagraphStyle('linea', array (
    'borderBottomSize' => 1,
    'borderColor' => '000000',
  ));

  $phpWord->addFontStyle('footer1',array(
    'color' => '8B0000',
    'bold' => true,
    'size' => 9 
  ));

  $phpWord->addFontStyle('footer2',array(
    'color' => '8B0000',
    'bold' => true,
    'italic' => true,
    'size' => 9 
  ));

    //BackGround -> 'bgColor' => '8B0000'

  $phpWord->addFontStyle('Asignatura',array(
    //'bgColor' => 'F0E68C',   Color de fondo
    'color' => '8B0000',
    'bold'=> true
  ));

  $phpWord->addParagraphStyle('AsigParrafo',array(
    'borderTopSize' => 1,
    'borderColor' => '000000'
  ));

  //poner borderTopSize para barra de arriba
  $phpWord->addParagraphStyle('FechaParrafo',array(
    'borderBottomSize' => 1,
    'borderColor' => '000000'
  ));

  $phpWord->addFontStyle('encabezados',array(
    'size' => 14,
    'color' => '8B0000',
    'bold' => true
  ));

  $phpWord->addParagraphStyle('encabezadoParrafo',array(
    'borderBottomSize' => 1,
    'borderColor' => '000000'
  ));

   $phpWord->addFontStyle('texto',array(
    'size' => 11,
    'color' => '000000',
    'bold' => false,
    'italic' => true
  ));
}

if ($_SERVER["REQUEST_METHOD"] == "POST") {

  if(sizeof($_SESSION) == 0 ){

    $_SESSION['nroUnidad'] = $_POST['nroUnidad'];
    $_SESSION['asignatura'] = $_POST['asignatura'];
    $_SESSION['fecha'] =  $_POST['fecha'];

    $_SESSION['cantO'] = $_SESSION['cantC'] = $_SESSION['cantD'] = 0;
    $_SESSION['cantA'] = $_SESSION['cantF'] = $_SESSION['cantB'] = 0;
    $_SESSION['cantS'] = 1;

    $_SESSION['objetivos'] = $_SESSION['contenidos'] = $_SESSION['desarrollo'] = array();
    $_SESSION['actividades'] = $_SESSION['foro'] = $_SESSION['bibliografia'] = array();

  }

  if(isset($_POST["CargarDatos"])){

    $_SESSION['nroUnidad'] = $_POST['nroUnidad'];
    $_SESSION['asignatura'] = $_POST['asignatura'];
    $_SESSION['fecha'] =  $_POST['fecha'];
  
  }

  else if(isset($_POST["seleccion"]) && isset($_POST['cargar'])){
    $seleccion = $_POST['seleccion'];

    switch ($seleccion) {
      case 'objetivos':
        $texto = $_POST["datos"];
        $cantO = $_SESSION['cantO'];
        $objetivos = $_SESSION['objetivos'];

        if(!empty($texto)){
          $objetivos[$cantO] = $texto;
          ++$cantO;
        }

        $_SESSION['cantO'] = $cantO;
        $_SESSION['objetivos'] = $objetivos;
        break;

      case 'contenidos':

        $texto = $_POST["datos"];
        $cantC = $_SESSION['cantC'];
        $contenidos = $_SESSION['contenidos'];

        if(!empty($texto)){
          $contenidos[$cantC] = $texto;
          ++$cantC;
        }

        $_SESSION['contenidos'] = $contenidos;
        $_SESSION['cantC'] = $cantC;
        break;

      case 'desarrollo':

        $texto = $_POST["datos"];
        $cantD = $_SESSION['cantD'];
        $desarrollo = $_SESSION['desarrollo'];

        if(!empty($texto)){
          $desarrollo[$cantD] = $texto;
          ++$cantD;
        }

        $_SESSION['desarrollo'] = $desarrollo;
        $_SESSION['cantD'] = $cantD;
        break;

      case 'actividades':

        $texto = $_POST["datos"];
        $cantA = $_SESSION['cantA'];
        $actividades = $_SESSION['actividades'];

        if(!empty($texto)){
          $actividades[$cantA] = $texto;
          ++$cantA;
        }

        $_SESSION['actividades'] = $actividades;
        $_SESSION['cantA'] = $cantA;
        break;

      case 'foro':

        $texto = $_POST["datos"];
        $cantF = $_SESSION['cantF'];
        $foro = $_SESSION['foro'];

        if(!empty($texto)){
          $foro[$cantF] = $texto;
          ++$cantF;
        }

        $_SESSION['foro'] = $foro;
        $_SESSION['cantF'] = $cantF;
        break;

      case 'bibliografia':

        $texto = $_POST["datos"];
        $cantB = $_SESSION['cantB'];
        $bibliografia = $_SESSION['bibliografia'];

        if(!empty($texto)){
          $bibliografia[$cantB] = $texto;
          ++$cantB;
        }

        $_SESSION['bibliografia'] = $bibliografia;
        $_SESSION['cantB'] = $cantB;
        break;
    }
  }
  else {

      $phpWord = new \PhpOffice\PhpWord\PhpWord();

      $seccion1 = $phpWord -> addSection(); 

      $header = $seccion1->addHeader();
      $header->addImage('encabezado.jpg',
      array(
        'wrappingStyle' => 'behind'
      ));

      $nro = $_SESSION['nroUnidad'];
      $asig = $_SESSION['asignatura'];
      $fecha = $_SESSION['fecha'];

      $cantS = $_SESSION['cantS'];
      $cantO = $_SESSION['cantO'];
      $cantC = $_SESSION['cantC'];
      $cantD = $_SESSION['cantD'];
      $cantA = $_SESSION['cantA'];
      $cantF = $_SESSION['cantF'];
      $cantB = $_SESSION['cantB'];

      $o = $c = $d = $a = $f = $b = 1;

      $seccion1->addText('GUÍA DE TRABAJO SEMANAL N° ' . $nro , 'titulo','titulo1');
      $seccion1->addText('','', 'linea');
      $seccion1->addText('Asignatura: '. $asig , 'Asignatura'); 
      $seccion1->addText('Año: ' . $fecha, 'Asignatura', 'FechaParrafo');
      $seccion1->addTextBreak();

      $objetivos = $_SESSION['objetivos'];
      $contenidos = $_SESSION['contenidos'];
      $desarrollo = $_SESSION['desarrollo'];
      $actividades = $_SESSION['actividades'];
      $foro = $_SESSION['foro'];
      $bibliografia = $_SESSION['bibliografia'];

      $seccion1->addText('OBJETIVOS','encabezados', 'encabezadoParrafo');

      if ($cantO>0){
            foreach ($objetivos as $obj) {
              $seccion1->addText($o . '.   ' .  $obj,'texto');
              ++$o;
            }
      }

      $seccion1->addTextBreak();
      $seccion1->addText('CONTENIDOS','encabezados', 'encabezadoParrafo');

      if ($cantC>0){
          foreach ($contenidos as $cont) {
            $seccion1->addText($c . '.   ' .  $cont,'texto');
            ++$c;
          }
      }

      $seccion1->addTextBreak();
      $seccion1->addText('DESARROLLO DE LA TEMÁTICA','encabezados', 'encabezadoParrafo');

      if ($cantD>0){
          foreach ($desarrollo as $des) {
            $seccion1->addText($d . '.   ' .  $des,'texto');
            ++$d;
          }
      }     

      $seccion1->addTextBreak();
      $seccion1->addText('ACTIVIDADES','encabezados', 'encabezadoParrafo');

      if ($cantA>0){
          foreach ($actividades as $act) {
            $seccion1->addText($a . '.   ' .  $act,'texto');
            ++$a;
          }
      } 

      $seccion1->addTextBreak();
      $seccion1->addText('FORO','encabezados', 'encabezadoParrafo');

      if ($cantF>0){
          foreach ($foro as $fo) {
            $seccion1->addText($f . '.   ' .  $fo,'texto');
            ++$f;
          }
      } 

      $seccion1->addTextBreak();
      $seccion1->addText('BIBLIOGRAFIA','encabezados', 'encabezadoParrafo');

      if ($cantB>0){
          foreach ($bibliografia as $bib) {
            $seccion1->addText($b . '.   ' .  $bib,'texto');
            ++$b;
          }
      } 

      $footer = $seccion1 -> createFooter();
      $footer->addText('GUÍA DE TRABAJO SEMANAL N° ' . $nro,'footer1');
      $footer->addText('Asignatura: ' . $asig, 'footer2');

      estilo($phpWord);
      $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
      session_unset();
      $objWriter->save('helloWorld.docx');
    }

}

?>

<!DOCTYPE HTML>
<html>  
<head>
  <meta charset="UTF-8">
  <link id="cssRef" href="estilo.css" type="text/css" rel="stylesheet">
</head>
<body>

<form id = "formulario" method="post">
    <img src = "encabezado.jpg" />
    <h2 id="formulario">Formulario para crear Guía de Trabajo Semanal</h2>
    Numero de Unidad: <input type="number" name="nroUnidad">
    <br><br> 
    Nombre de Unidad: <input type="text" name="nombreUnidad">
    <br><br> 
    Asignatura: <input type="text" name="asignatura">
    <br><br> 
    Fecha: <input type=date name = "fecha">
    <br><br> 
    <input id="boton1" type="submit" name="CargarDatos" value="Cargar Datos">  
    <br><br> 
    Seleccione el dato que desea cargar:
    <br><br>
    <select name="seleccion">
      <option value="objetivos">Objetivos</option>
      <option value="contenidos">Contenidos</option>
      <option value="desarrollo">Desarrollo de la temática</option>
      <option value="actividades">Actividades</option>
      <option value="foro">Foro</option>
      <option value="bibliografia">Bibliografia</option>
    </select>
    <br><br> 
    <textarea name="datos" rows="3" cols="30"></textarea>
    <br><br>
    <input id ="boton2" type="submit" name="cargar" value="Cargar Item"/> 
    <br><br>
    Para crear el word final:
    <input id ="boton3" type="submit" name="crearWord" value="Crear Word">  
    <br><br> 
    Para crear el pdf final
    <input id ="boton4" type="submit" name="crearPDF" value="Crear PDF">  

</form>

</body>
</html>

