<?php
/**
 * Demostrar lectura de hoja de cálculo o archivo
 * de Excel con PHPSpreadSheet: leer todo el contenido
 * de un archivo de Excel usando índices, no iteradores
 *  MODIFICADO POR JORGE EMILIANO MALDONADO
 * @author parzibyte
 */
# Cargar librerias y cosas necesarias
require_once "vendor/autoload.php";

# Indicar que usaremos el IOFactory
use PhpOffice\PhpSpreadsheet\IOFactory;

# Recomiendo poner la ruta absoluta si no está junto al script
# Nota: no necesariamente tiene que tener la extensión XLSX
// $rutaArchivo = "LibroParaLeerConPHP.xlsx";
$rutaArchivo = "datos.xls";

$documento = IOFactory::load($rutaArchivo);

# Recuerda que un documento puede tener múltiples hojas
# obtener conteo e iterar
$totalDeHojas = $documento->getSheetCount();

$matriz = [];


# Iterar hoja por hoja
for ($indiceHoja = 0; $indiceHoja < $totalDeHojas; $indiceHoja++) {
    # Obtener hoja en el índice que vaya del ciclo
    $hojaActual = $documento->getSheet($indiceHoja);
    echo "<h3>índice $indiceHoja</h3>";

    # Calcular el máximo valor de la fila como entero, es decir, el
    # límite de nuestro ciclo
    $numeroMayorDeFila = $hojaActual->getHighestRow(); // Numérico
    $letraMayorDeColumna = $hojaActual->getHighestColumn(); // Letra
    # Convertir la letra al número de columna correspondiente
    $numeroMayorDeColumna = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($letraMayorDeColumna);


    echo '<table border="1">';
    $dato = ''; 
    # Iterar filas con ciclo for e índices
    for ($indiceFila = 1; $indiceFila <= $numeroMayorDeFila; $indiceFila++) {
        $Y = $indiceFila;
        echo '<tr>';
        for ($indiceColumna = 1; $indiceColumna <= $numeroMayorDeColumna; $indiceColumna++) {
            $X = $indiceColumna;
            # Obtener celda por columna y fila
            $celda = $hojaActual->getCellByColumnAndRow($X, $Y);
            # Y ahora que tenemos una celda trabajamos con ella igual que antes
            # El valor, así como está en el documento
            $valorRaw = $celda->getValue();

            # Formateado por ejemplo como dinero o con decimales
            $valorFormateado = $celda->getFormattedValue();

            # Si es una fórmula y necesitamos su valor, llamamos a:
            $valorCalculado = $celda->getCalculatedValue();

            # Fila, que comienza en 1, luego 2 y así...
            $fila = $celda->getRow();
            # Columna, que es la A, B, C y así...
            $columna = $celda->getColumn();

            // $dato = $hojaActual->getCellByColumnAndRow(0, 0);
            
            echo '<th>';
            
            #echo "$valorRaw";
            echo "$valorFormateado";
            #echo "Calculado es: <strong>$valorCalculado</strong><br><br>";
            echo '</th>';
        }
        echo '</tr>';
    }

    // var_dump($dato);
    // echo '<hr>';


    echo '</table>';

}
