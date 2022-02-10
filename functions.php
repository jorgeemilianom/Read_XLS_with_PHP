<?php
/**
 * Libreria editada por Jorge Emiliano Maldonado
 * 
 *
 * @author parzibyte
 */

require_once "vendor/autoload.php";
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpParser\Node\Stmt\Return_;

function readAllDocument($fileRute, $nHoja){
    $rutaArchivo = $fileRute;
    $documento = IOFactory::load($rutaArchivo);
    $totalDeHojas = $documento->getSheetCount();
    $matriz = [];

    $hojaActual = $documento->getSheet($nHoja);

    $numeroMayorDeFila = $hojaActual->getHighestRow();
    $letraMayorDeColumna = $hojaActual->getHighestColumn();
    $numeroMayorDeColumna = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($letraMayorDeColumna);

    for ($indiceFila = 1; $indiceFila <= $numeroMayorDeFila; $indiceFila++) {
        $Y = $indiceFila;
        for ($indiceColumna = 1; $indiceColumna <= $numeroMayorDeColumna; $indiceColumna++) {
            $X = $indiceColumna;
            $celda = $hojaActual->getCellByColumnAndRow($X, $Y);
            $valorRaw = $celda->getValue();
            $valorFormateado = $celda->getFormattedValue();
            $valorCalculado = $celda->getCalculatedValue();
            $fila = $celda->getRow();
            $columna = $celda->getColumn();
            if(!empty($valorFormateado)){
                $matriz[$Y][$X] = $valorFormateado;
            }
        }
    }
    return $matriz;
}

function readDocumentLimits($fileRute, $nHoja, $coordenate1, $coordenate2){    // 1er Point    2do Point
    $limX = [$coordenate1[0], $coordenate2[0]];
    $limY = [$coordenate1[1], $coordenate2[1]];
    $rutaArchivo = $fileRute;
    $documento = IOFactory::load($rutaArchivo);
    $totalDeHojas = $documento->getSheetCount();
    $matriz = [];

    $hojaActual = $documento->getSheet($nHoja);

    $numeroMayorDeFila = $hojaActual->getHighestRow();
    $letraMayorDeColumna = $hojaActual->getHighestColumn();
    $numeroMayorDeColumna = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($letraMayorDeColumna);

    for ($indiceFila = $limY[0]; $indiceFila <= $limY[1]; $indiceFila++) {
        $Y = $indiceFila;
        for ($indiceColumna = letterWithNumber($limX[0]); $indiceColumna <= letterWithNumber($limX[1]); $indiceColumna++) {
            $X = $indiceColumna;
            $celda = $hojaActual->getCellByColumnAndRow($X, $Y);
            $valorRaw = $celda->getValue();
            $valorFormateado = $celda->getFormattedValue();
            $valorCalculado = $celda->getCalculatedValue();
            $fila = $celda->getRow();
            $columna = $celda->getColumn();
            if(!empty($valorRaw)){
                $matriz[$Y][$X] = $valorRaw;
            }else{
                $matriz[$Y][$X] = '-';
            }
        }
    }
    return $matriz;
}

function letterWithNumber($let){
    $letters = "ABCDEFGHIJKLMNOPQRSTUVWHYZ";
    $lettersMin = "abcdefghijklmnopqrstuvwxyz";
    $n = 0;
    for($i=0; $i<strlen($let); $i++){
        if($i+1<strlen($let)){
            $n+=25;
        }
        for($x = 0; $x<strlen($letters); $x++){
            if($let[$i] == $letters[$x] || $let[$i] == $lettersMin[$x]){
                $n += $x+1;
            }
        }
    }
    return $n;
}
