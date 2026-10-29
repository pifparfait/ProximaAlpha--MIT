# Entorno local — figuras IJGFS-D-26-01381

## Estructura esperada

```
proyecto/
├── generate_figures.py   ← este script
├── requirements.txt
├── data/                 ← poner aquí los 4 Excel del repo
│   ├── movement1_pig_duck_results.xlsx
│   ├── movement2_aroma_results_v2.xlsx
│   ├── movement3_pde_results_<hash>.xlsx
│   └── movement4_optimization_results.xlsx
└── figures/              ← se crea automáticamente
    ├── f1.pdf / f1.png
    ├── f3.pdf / f3.png
    ...
```

## Setup (una sola vez)

```bash
# Python 3.9+ recomendado
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Generar figuras

```bash
python generate_figures.py --data ./data --out ./figures
```

## Correspondencia figura ↔ nombre LaTeX

| Archivo generado | Figura en el paper       |
|------------------|--------------------------|
| f1.pdf/png       | Fig 1: Layer geometry    |
| f3.pdf/png       | Fig 2: Drying kinetics   |
| f5.pdf/png       | Fig 3: Aromatic pathways |
| f6.pdf/png       | Fig 4: Reaction rates    |
| f10.pdf/png      | Fig 5: Temperature evol  |
| f12.pdf/png      | Fig 6: PDE spatial fields|
| f14.pdf/png      | Fig 7: Pareto + regression|

El LaTeX los incluye así: `\includegraphics[...]{f1}`, `{f3}`, etc.
Pon los PDF en la misma carpeta que el .tex al compilar.

## Ver figuras con ventana (opcional)

Comenta la línea `matplotlib.use("Agg")` al inicio del script.

## Datos fuente

Los Excel están en el repositorio público:
  https://github.com/blackmontAI/ProximaAlpha--MIT
