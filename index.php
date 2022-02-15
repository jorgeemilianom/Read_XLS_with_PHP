<?php
/**
 * Esta libreria es una mejora de https://github.com/parzibyte/leer_excel_php
 * @author Jorge Emiliano Maldonado
 */

require('functions.php');

// $allData = readAllDocument('datos.xls', 1);
$data = readDocumentLimits(
    'datos.xlsx',       // File XLS (EXCEL)
    0,                  // Sheet
    ['A', 1],           // Starting point
    ['F', 7]            // End Point
);

$dataJSON = json_encode($data);
echo($dataJSON);


?>