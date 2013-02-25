<?php
/*
Código fuente: classes.php
Descrición: Clases para importar desde una tabla de una base de datos MySQL a una planilla XLS.
Version: 1.0
Autor: Lucas Dima
email: lucdima@gmail.com
Licencia: GNU/GPLv3

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.


Modo de uso:
La clase Conector sirve para establecer la conección con la base de datos. Por default utiliza UTF-8.
La clase Excel es la que se encarga de generar el archivo XLS.
En la línea 36 se puede cambiar la consulta en caso de ser necesario.
*/

class Conector {
		function abrirbase($server,$database,$dbuser,$pass) {
			$link=mysql_connect($server,$dbuser,$pass);
			if (!$link)
				die("<br/>No se puede conectar a la base de datos");
				mysql_set_charset("utf8",$link);
				mysql_select_db($database,$link)
			or die ("<br/>No se puede abrir $database:".mysql_error());
			return $link;
		}
}


class Excel {
	
	private $conector;
	
	function generar($server,$database,$dbuser,$pass,$table,$nombreArchivo) {
		$c=new Conector();
		$c->abrirbase($server,$database,$dbuser,$pass);
		
		$querytxt="SELECT * FROM ".$table.";";
		$result=mysql_query($querytxt);

		if (mysql_error()) {
			die("<br/>Error al acceder a la tabla: ".mysql_error());
		}

		if (mysql_num_rows($result)==0) {
			die("<br/>No hay datos en la tabla");
		}
	
	
	// Functions for export to excel.
		function xlsBOF() {
		echo pack("ssssss", 0x809, 0x8, 0x0, 0x10, 0x0, 0x0);
		return;
		}
		function xlsEOF() {
			echo pack("ss", 0x0A, 0x00);
			return;
		}
	function xlsWriteNumber($Row, $Col, $Value) {
		echo pack("sssss", 0x203, 14, $Row, $Col, 0x0);
		echo pack("d", $Value);
	return;
}
	function xlsWriteLabel($Row, $Col, $Value ) {
		$L = strlen($Value);
		echo pack("ssssss", 0x204, 8 + $L, $Row, $Col, 0x0, $L);
		echo $Value;
		return;
	}


header("Pragma: public");
header("Expires: 0");
header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
header("Content-Type: application/force-download");
header("Content-Type: application/octet-stream");
header("Content-Type: application/download");;
header("Content-Disposition: attachment;filename=".$nombreArchivo);
header("Content-Transfer-Encoding: binary ");

xlsBOF();

$i = 0;
while ($i < mysql_num_fields($result)) {
    $meta = mysql_fetch_field($result, $i);
    if (!$meta) {
        echo "No hay datos\n";
    }
	xlsWriteLabel(0,$i,$meta->name);
	$fieldnames[$i]=$meta->name;
	$fieldtype[$i]=$meta->numeric;
    $i++;
}

$xlsRow = 1;
$j=0;

while($row=mysql_fetch_array($result)){
	for ($j=0;$j<$i;$j++) {
		if ($fieldtype[$j]==0) {
			xlsWriteLabel($xlsRow,$j,$row[$fieldnames[$j]]);
		}
		else {
			xlsWriteNumber($xlsRow,$j,$row[$fieldnames[$j]]);
		}
	}
	
$xlsRow++;
}
xlsEOF();
}
}
?>
