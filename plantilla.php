<?php

require 'vendor/autoload.php';

use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\Settings;


use PhpOffice\PhpWord\Element\TextRun;

error_reporting(E_ALL);

//hacer una copia del archivo plantilla.docx con el nombre 'plantillla_'.date('YmdHmi').'.docx'
//y guardarlo en una carpeta temporal

$ruta_plantilla = 'plantilla/plantilla.docx';
$ruta_archivo = 'plantilla.docx';
$ruta_archivo_copia = 'plantilla_'.date('YmdHis').'.docx';
$ruta_carpeta_temporal = 'tmp/';

//BORRADO DE ARCHIVOS TEMPORALES
$ficheros = scandir($ruta_carpeta_temporal,0);
foreach($ficheros as $file){
    if ($file > '..'){
        unlink($ruta_carpeta_temporal.$file);
    }
}

$doc_path = $ruta_carpeta_temporal.$ruta_archivo_copia;

if (copy($ruta_plantilla, $doc_path)) {
    echo 'Archivo copiado<br>';
} else {
    echo 'Error al copiar';
    die();
}

$template = new \PhpOffice\PhpWord\TemplateProcessor($doc_path);

$detalles = [
        'titulo' => 'Titulo de la plantilla º',
        // 'RequisitoSanitarioPais' => ['Texto de la plantilla','Texto de la plantilla','Texto de la plantilla'],
        'otro_texto' => 'Otro texto de la plantilla',
		'col_1' => [
			[
				'col_1' => 'lala',
				'col_2' => 'lala',
				'col_3' => 'lala'
			],
			[
				'col_1' => 'xxx',
				'col_2' => 'xxa',
				'col_3' => 'laxx'
			]
		]
];

$style = [
        // 'italic' => true, 
        // 'color' => 'red',
        // 'name' => 'Tahoma', 
        // 'size' => 10, 
        'strikethrough' => true
    ];

foreach ($detalles as $key => $value){
    
    if (is_array($value)){
        $template->cloneRow($key, count($value));
        foreach ($value as $idx => $arr_inside){

			foreach($arr_inside as $k => $v){
				$template->setValue($k.'#'.($idx+1), $v);				
			}
            //$inline = new TextRun();
            //$inline->addText($texto, $style);
            //$template->setComplexValue($key.'#'.($idx+1), $inline);
            
        }
    }else{
        $template->setValue($key, $value);
    }
}
$template->saveAs($doc_path);

//TODO TEST THIS
//$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'HTML'); 

// Make sure you have `dompdf/dompdf` in your composer dependencies.
Settings::setPdfRendererName(Settings::PDF_RENDERER_DOMPDF);
// Any writable directory here. It will be ignored.
Settings::setPdfRendererPath('.');

$phpWord = IOFactory::load($doc_path, 'Word2007');
$phpWord->save('document_x.pdf', 'PDF');