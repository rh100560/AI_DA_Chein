# Análisis de clientes de delivery de comida

Proyecto de análisis exploratorio y clasificación con un dataset de clientes de una plataforma de delivery en Bengaluru, India. El trabajo explora asociaciones entre características de los clientes y construye un modelo exploratorio para predecir si volverán a ordenar.

## Contenido

- `notebooks/01_exploracion_dataset.ipynb`: exploración inicial del dataset.
- `notebooks/insight_*.ipynb`: análisis de relaciones seleccionadas.
- `notebooks/modelo_recompra_output.ipynb`: entrenamiento, validación, test e importancia de variables.
- `presentacion_final_delivery.pptx`: presentación final con hallazgos, gráficas y resultados del modelo.
- `scripts/crear_presentacion.py`: genera la presentación a partir de los datos.
- `data/online food delivery dataset.csv`: datos utilizados en el proyecto.

## Pregunta de análisis

¿Se puede predecir si un cliente volverá a ordenar (`Output`) a partir de su perfil y ubicación?

Es un problema de clasificación binaria. `Yes` significa que volverá a ordenar y `No` que no volverá. Los predictores utilizados son edad, género, estado civil, ocupación, ingreso mensual, educación, tamaño familiar, latitud y longitud.

`Feedback` se excluye porque podría registrarse después del pedido y revelar información futura. También se excluyen `Customer Type`, que replica `Family size` en este dataset, y `Pin code`, que es un identificador geográfico.

## Dataset y preparación

- 388 filas originales y 13 variables después de quitar una columna vacía que se verificó como copia exacta de `Output`.
- 103 filas repetidas eliminadas; quedan 285 registros únicos.
- Clases del objetivo: 217 `Yes` (76.1%) y 68 `No` (23.9%).
- División estratificada: 171 registros de entrenamiento, 57 de validación y 57 de test.

## Hallazgos exploratorios

La fuerza de asociación entre variables categóricas se midió con V de Cramer corregida. Los valores indican asociación, no causalidad.

| Relación | V de Cramer corregida | Lectura |
| --- | ---: | --- |
| Tamaño familiar y tipo de cliente | 0.995 | En estos datos, los valores siguen una correspondencia casi determinista; podría ser una regla de codificación. |
| Resultado y feedback | 0.545 | Asociación alta, aunque ambas variables describen aspectos relacionados de la experiencia. |
| Ocupación e ingreso mensual | 0.524 | Los rangos de ingreso varían claramente según la ocupación. |
| Educación y ocupación | 0.373 | Se observa una asociación moderada entre nivel educativo y ocupación. |

## Modelado y resultados

Se compararon un baseline que predice la clase más frecuente, regresión logística y random forest. Los modelos se ajustaron en entrenamiento y se eligieron con validación. La selección prioriza balanced accuracy y usa F1 para `No` como desempate. El test se reservó para evaluar una sola vez el modelo elegido.

| Modelo | Balanced accuracy (train) | Balanced accuracy (validación) | F1 `No` (validación) | Resultado en test |
| --- | ---: | ---: | ---: | --- |
| Baseline (clase frecuente) | 0.500 | 0.500 | 0.000 | — |
| Regresión logística | 0.674 | 0.660 | 0.489 | — |
| Random forest (seleccionado) | 1.000 | 0.669 | 0.500 | Accuracy: 0.684; balanced accuracy: 0.670; F1 `No`: 0.500 |

El random forest detectó 9 de los 14 casos `No` del test y clasificó correctamente 30 de 43 casos `Yes`. Su balanced accuracy en entrenamiento (1.000) es mucho mayor que en validación (0.669), lo que sugiere sobreajuste.

En la importancia por permutación sobre validación destacaron latitud, edad y ocupación. La muestra de validación es pequeña y las estimaciones tienen incertidumbre; no deben interpretarse como efectos causales ni como un ranking definitivo.

## Conclusión y limitaciones

El modelo supera al baseline al detectar ambas clases, pero su rendimiento es exploratorio y la brecha entre entrenamiento y validación apunta a sobreajuste. El test contiene solo 14 casos `No`, así que unos pocos aciertos o errores modifican considerablemente las métricas. Para evaluar un uso real hacen falta más observaciones independientes, validación cruzada estratificada y confirmar que las variables estén disponibles antes de predecir.

## Cómo reproducir

Requiere Python 3.12 o superior. Desde la raíz del proyecto:

```bash
uv sync
```

Abrir y ejecutar `notebooks/modelo_recompra_output.ipynb` de arriba a abajo. Para regenerar las diapositivas:

```bash
uv run python scripts/crear_presentacion.py
```

## Herramientas

Python, pandas, scikit-learn, SciPy, matplotlib, seaborn y python-pptx.

## Autores

Regina Hernández y Matias Cohen
