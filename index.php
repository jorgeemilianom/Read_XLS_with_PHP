<?php
/**
 * Esta libreria es una mejora de https://github.com/parzibyte/leer_excel_php
 * @author Jorge Emiliano Maldonado
 */

require('functions.php');

//  readDocumentLimits('RuteFile', nShits, Coordenate1['A', 4], Coordenate2['F', 15])

// $data = readAllDocument('datos.xls', 1);
$data = readDocumentLimits('datos.xlsx', 0, ['C', 6], ['AD', 36]);

$dataJSON = json_encode($data);
echo($dataJSON);


?>