<?php
/*
Código fuente: exportxls.php
Descripción: Importa desde una tabla de una base de datos MySQL a una planilla XLS.
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
Reemplace el valor de las varialbes $server,$database,$databaseuser,$databasepassword,$table,$filename por el de su servidor (default:localhost), base de datos, usuario de la base de datos, password del usuario, tabla y nombre de archivo respectivamente.
*/
include ('classes.php');

// Especificación del servidor.
$server="localhost";
$database="mibasededatos";
$databaseuser="usuario";
$databasepassword="miclave";
$table="mitabla";
$filename="nombre.xls";

// Generación de la planilla
$e = new Excel();
$e->generar($server,$database,$databaseuser,$databasepassword,$table,$filename);
?>