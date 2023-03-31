# docReader

Script para leer un párrafo específico y tablas dentro de un documento .docx

Dentro del script, en la línea 21 está la ruta del archivo a ser leído:

```python
file = "C:\\Users\\lorni\\Documents\\Desarrollo\\prueba_lectura_archivos_doc\\informe_unix_con_comentarios2.docx"
```

El script va a buscar ese archivo, y una vez encontrado buscará un párrafo (linea 9) y las tablas que contenga el archivo

La idea es atribuirles a cada tabla una propiedad nombre que se use en el script para luego identificar cada tabla con ese nombre dentro del archivo json resultante.

Ese archivo json resultante es el que se va a leer y a convertir en otro formato o pasar a una BD. Lo que convenga.

