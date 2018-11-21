Funciones-Excel
===============

Grupo de funciones de Excel en VBA para análisis de datos de encuestas y muestras.

Archivo: **`errores.bas`** contiene las siguientes funciones:

- `errormuestralinf`: Calcula el error muestral de para una muestra aleatoria simple obtenida de una población infinita (mayor a 100.000 elementos)
    
	Sus variables son:

    * `muestra`: Tamaño de la muestra.
    * `signif`: nivel de confianza. Por defecto `signif = 0.95`.
    * `p`: varianza. Por defecto `p = 0.5`.

- `errormuestralfin`: Calcula el error muestral de para una muestra aleatoria simple obtenida de una población finita.

	Sus variables son:

    * `muestra`: Tamaño de la muestra.
    * `pobTot`: Tamaño de la población.
    * `signif`: nivel de confianza. Por defecto `signif = 0.95`.
    * `p`: varianza. Por defecto `p = 0.5`.

- `tamuestra`: Calcula el tamaño muestral necesario para un error muestral dado.

	Sus variables son:

    * `error`: error muestral.
    * `pobTot`: Tamaño de la población.
    * `signif`: nivel de confianza. Por defecto `signif = 0.95`.
    * `p`: varianza. Por defecto `p = 0.5`.

- `errormuestral_dist`: Calcula el error muestral de para una muestra aleatoria obtenida de una población finita.

	Sus variables son:

    * `muestra`: Rango de estratos muestrales.
    * `pob`: Rango de población según estratos muestrales.
    * `signif`: nivel de confianza. Por defecto `signif = 0.95`.
    * `p`: varianza. Por defecto `p = 0.5`.

Archivo **`luhn.bas`** contiene funciones para la creación y validación de números bajo el [algoritmo de luhn](https://en.wikipedia.org/wiki/Luhn_algorithm).
