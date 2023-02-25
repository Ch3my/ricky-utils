# Reporte de Asistencia
Este programa (funciona en Linux y Windows) toma el archivo de reporte de asistencia 
y calcula las horas extras y las horas nocturnas
con las reglas que se detallan mas adelante.

Para funcionar, el programa se alimenta de datos en una hoja llamada "Personas" donde esta el nombre
de la persona en la Columna A (tal cual se encuentra en el Excel), la cantidad de minutos
que esa persona trabaja en la columna B y la hora de entrada teorica (hora oficial de ingreso).
Con estos datos calcula las horas extra si corresponden.

Luego recorre todas las filas de la Hoja llamada "Asistencia" y revisa si el nombre de la fila
corresponde a algun nombre en la Hoja "Personas" sino corresponde se lo salta.

Luego revisa que existen los datos necesarios para hacer los calculos. Si viene algun dato
vacio (Error de informacion) avisa error por la consola y marca la fila en rojo.

Si procesa una fila (cumple condiciones para horas extra o nocturnas) luego de llenar
los datos que correspondan marca la fila en amarillo (quiza se comente en el futuro)
para facil revision del encargado

## Horas Extra
Para que cuenten horas extra la persona debe trabajar mas de los minutos que se especificaron
en la hoja "Personas". Ademas, debe haber trabajado mas de 20min extra, cualquier cosa menor a 20min
se considerara que es un retraso en la marcacion solamente y no horas extra.

Las horas extra se pueden contar desde antes del horario de ingreso o despues. Para que cuenten antes
del horario de ingreso la persona debe ingresar por lo menos 30min antes. Si llega menos de 30min antes
se considerara que ingreso a su horario teorico (horario oficial de ingreso)

## Horas nocturnas
Las horas nocturnas cuentan desde las 00:00:00 hasta las 08:00:00. Por lo tanto una persona puede
tener como maximo 8 horas nocturnas.

Para determinar que son horas nocturas la persona debio marcar ingreso despues de las 17:00 y salir
despues de las 00:00:00.

## Control de Errores
Si hay una cantidad exagerada de horas extras u horas nocturnas el sistema asume que fue un error de marcacion
y marca la fila con rojo para revision del encargado

## Requerimientos de la informacion base
- Python3. Instalar dependencias `pip install openpyxl` `pip install printy`
- Debe existir una hoja "Personas" con 3 columnas segun se detallo anteriormente.
- El proceso de la informacion siempre se hace desde la segunda fila hacia abajo
- El formato de las horas en la hoja "Personas" es 23:59
- El formato de la fecha en la hoja "Asistencia" es dd/mm/yyyy
- El formato de la hora en la hoja "Asistencia" es 23:59:59
- La columna donde se escribe la informacion de horas extra es "AC"
- La columna donde se escribe la informacion de horas nocturas es "AD"