# addMailBox
Script para añadir un nuevo buzón en nuestro servidor de Exchange.

Este script nos permitirá:
* Añadir un nuevo buzón de correo dentro de nuestro servidor Exchange.
* Añadir buzones de manera masiva, usando CSV, a nuestro servidor Exchange. (En construcción).

¿Cómo funciona?
======
Este script tiene un menú donde nos va a pedir que opción queremos hacer (1-Para crear un buzón de manera individual, 2-Crear buzones de manera masiva). La opción 2 aun está en fase de desarrollo, por lo tanto, no está activa.
La opción 1, nos va a pedir los datos del buzón; nombre, apellido, dirección de correo y finalmente el Tenant(OU) donde tenemos alojado este cliente. El propio script realiza las comprobaciones para saber que todo lo que hemos indicado es correcto, si el usuario no existe dentro del Tenant nos creará el buzón de correo usando una contraseña randomizada de 16 caracteres.
Una vez terminado el proceso, por pantalla nos indicará el resumen de la operación y nos indicará la contraseña que se ha creado para este buzón.

Posibles utilidades
======
* Creación de forma automatizada de buzones de correo Exchange en un servidor propio.

Uso
======
Editar la variable PATH dentro del script para indicarle los datos, por ejemplo dominio, OU, etc.
Una vez cambiado el PATH ejecutar el script y seguir los pasos que indica por pantalla.
