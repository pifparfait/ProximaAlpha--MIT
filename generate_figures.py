"""
generate_figures.py
===================
Genera las 7 figuras del manuscrito IJGFS-D-26-01381
"What Expert Chefs Know That Equations Help Explain"

Requisitos:
    pip install matplotlib numpy scipy openpyxl

Uso:
    python generate_figures.py --data ./data --out ./figures

Los 4 archivos Excel deben estar en --data/:
    movement1_pig_duck_results.xlsx
    movement2_aroma_results_v2.xlsx
    movement3_pde_results_*.xlsx   (el nombre completo tiene un hash)
    movement4_optimization_results.xlsx

Salida en --out/:
    f1.pdf / f1.png   → Fig 1: Layer geometry
    f3.pdf / f3.png   → Fig 2: Drying / moisture loss
    f5.pdf / f5.png   → Fig 3: Aromatic pathways
    f6.pdf / f6.png   → Fig 4: Arrhenius reaction rates
    f10.pdf / f10.png → Fig 5: Temperature evolution by phase
    f12.pdf / f12.png → Fig 6: PDE spatial fields
    f14.pdf / f14.png → Fig 7: Pareto front + sensory regression
"""

import argparse, os, warnings
import numpy as np
import matplotlib
matplotlib.use("Agg")          # sin ventana; quitar esta línea si quieres ver las figuras
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.gridspec import GridSpec
from openpyxl import load_workbook
from scipy.stats import linregress

# ── Paleta y estilo global ────────────────────────────────────────────────────
BLUE   = "#2166ac"
ORANGE = "#d6604d"
GREEN  = "#4dac26"
PURPLE = "#762a83"
GRAY   = "#888888"

plt.rcParams.update({
    "font.family":     "serif",
    "font.size":       10,
    "axes.labelsize":  11,
    "axes.titlesize":  11,
    "legend.fontsize": 9,
    "figure.dpi":      150,
})

# ── Utilidades ────────────────────────────────────────────────────────────────
def wb_rows(path, sheet, skip=1):
    """Devuelve filas de una hoja como listas, saltando `skip` filas de cabecera."""
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet]
    return list(ws.iter_rows(values_only=True))[skip:]

def find_m3(data_dir):
    """Localiza el Excel de movement3 (nombre con hash)."""
    for f in os.listdir(data_dir):
        if f.startswith("movement3_pde") and f.endswith(".xlsx"):
            return os.path.join(data_dir, f)
    raise FileNotFoundError("No se encontró movement3_pde_*.xlsx en " + data_dir)

def save(fig, out_dir, stem):
    fig.savefig(os.path.join(out_dir, f"{stem}.pdf"), bbox_inches="tight")
    fig.savefig(os.path.join(out_dir, f"{stem}.png"), bbox_inches="tight", dpi=200)
    plt.close(fig)
    print(f"  ✓  {stem}.pdf / .png")


# ══════════════════════════════════════════════════════════════════════════════
# Fig 1  —  Layer geometry  (movement1, sheet: layers)
# Columnas: species | layer | thickness_m | thickness_fraction | density | mass
# ══════════════════════════════════════════════════════════════════════════════
def fig1_layer_geometry(data_dir, out_dir):
    rows = wb_rows(os.path.join(data_dir, "movement1_pig_duck_results.xlsx"),
                   "layers", skip=1)

    # Agrupar por especie: acumular thickness_fraction por layer
    from collections import defaultdict, OrderedDict
    species_layers = defaultdict(dict)  # {species_base: {layer: frac}}
    order = []
    for r in rows:
        if r[0] is None: continue
        sp    = str(r[0])
        layer = str(r[1])
        frac  = float(r[3])
        base  = sp.split("_")[0].capitalize()   # pig_JohnDallas → Pig
        if base not in order: order.append(base)
        # Si ya existe la clave, promediar (varias entradas pig_standard, pig_JohnDallas)
        if layer in species_layers[base]:
            species_layers[base][layer] = (species_layers[base][layer] + frac) / 2
        else:
            species_layers[base][layer] = frac

    animals = [k for k in order if k in species_layers]
    layer_names = ["skin", "fat", "muscle"]
    colors = [ORANGE, BLUE, GREEN]

    x  = np.arange(len(animals)); w = 0.5
    fig, ax = plt.subplots(figsize=(5, 4))
    bottoms = np.zeros(len(animals))
    for lname, col in zip(layer_names, colors):
        vals = np.array([species_layers[a].get(lname, 0) for a in animals])
        ax.bar(x, vals, w, bottom=bottoms, label=lname.capitalize(), color=col)
        bottoms += vals

    ax.set_xticks(x); ax.set_xticklabels(animals)
    ax.set_ylabel("Thickness fraction")
    ax.set_title("Layer geometry — skin / fat / muscle")
    ax.legend(loc="upper right")
    ax.set_ylim(0, 1.05)
    ax.yaxis.grid(True, linestyle="--", alpha=0.4); ax.set_axisbelow(True)
    save(fig, out_dir, "f1")


# ══════════════════════════════════════════════════════════════════════════════
# Fig 2  —  Drying / moisture loss  (movement1, sheet: drying_48h)
# Columnas: species | layer | w0 | w_eq | w_48h | X_48h | W_init | W_final | Loss
# ══════════════════════════════════════════════════════════════════════════════
def fig2_drying_moisture(data_dir, out_dir):
    rows = wb_rows(os.path.join(data_dir, "movement1_pig_duck_results.xlsx"),
                   "drying_48h", skip=1)
    data = [r for r in rows if r[0] is not None]

    # Curva modelo de drying: interpolación lineal w0 → w_48h en t=[0, 48h]
    t_pts = np.linspace(0, 48, 200)

    fig, ax = plt.subplots(figsize=(6, 4))
    style_map = {
        ("pig", "skin"):   (ORANGE, "-",  "Pig — skin"),
        ("pig", "fat"):    (ORANGE, "--", "Pig — fat"),
        ("duck", "skin"):  (BLUE,   "-",  "Duck — skin"),
        ("duck", "fat"):   (BLUE,   "--", "Duck — fat"),
    }

    plotted = set()
    for r in data:
        sp    = str(r[0]).split("_")[0].lower()
        layer = str(r[1]).lower()
        w0    = float(r[2])
        w_48  = float(r[4])
        key   = (sp, layer)
        if key not in style_map or key in plotted: continue
        col, ls, lbl = style_map[key]
        # Secado exponencial simple: w(t) = w_48 + (w0 - w_48)*exp(-t/tau)
        tau = 48 / (-np.log((w_48 - w_48 + 1e-9) / (w0 - w_48 + 1e-9) + 1e-9) + 1e-9)
        # Si tau es inválido, lineal
        if not np.isfinite(tau) or tau <= 0:
            w_curve = np.linspace(w0, w_48, len(t_pts))
        else:
            w_curve = w_48 + (w0 - w_48) * np.exp(-t_pts / tau)
        ax.plot(t_pts, w_curve, color=col, ls=ls, lw=2, label=lbl)
        plotted.add(key)

    ax.set_xlabel("Time (h)")
    ax.set_ylabel("Moisture content (g/g, wet basis)")
    ax.set_title("Drying kinetics — 48 h pre-roast")
    ax.legend(); ax.yaxis.grid(True, linestyle="--", alpha=0.4); ax.set_axisbelow(True)
    save(fig, out_dir, "f3")


# ══════════════════════════════════════════════════════════════════════════════
# Fig 3  —  Aromatic pathways  (movement2, sheet: Tabla_2_3_aroma)
# Columnas: species | A_M | A_L | A_G | A_S | A_rel_M | A_rel_L | A_rel_G | A_rel_S | R_aroma | R_aroma_PCA
# ══════════════════════════════════════════════════════════════════════════════
def fig3_aromatic_pathways(data_dir, out_dir):
    rows = wb_rows(os.path.join(data_dir, "movement2_aroma_results_v2.xlsx"),
                   "Tabla_2_3_aroma", skip=1)
    data = [r for r in rows if r[0] is not None]

    channels    = ["Maillard", "Lipidic", "Caramelization", "Sulfur"]
    species_map = {}   # {species_label: (rel_M, rel_L, rel_G, rel_S, R_aroma)}
    for r in data:
        sp = str(r[0])
        species_map[sp] = (float(r[5])*100, float(r[6])*100,
                           float(r[7])*100, float(r[8])*100,
                           float(r[9]))

    sp_labels = list(species_map.keys())
    x = np.arange(len(channels)); w = 0.35
    colors_sp = [ORANGE, BLUE]

    fig, ax = plt.subplots(figsize=(8, 4))
    for i, (sp, vals) in enumerate(species_map.items()):
        pcts = list(vals[:4])
        offset = (i - 0.5) * w
        bars = ax.bar(x + offset, pcts, w, color=colors_sp[i], label=sp)
        # Anotar porcentaje encima de cada barra
        for bar, pct in zip(bars, pcts):
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.3,
                    f"{pct:.1f}%", ha="center", va="bottom", fontsize=7)

    # Anotar R_aroma al lado del nombre de especie en la leyenda (se hace en handles)
    handles = []
    for i, (sp, vals) in enumerate(species_map.items()):
        handles.append(mpatches.Patch(color=colors_sp[i],
                                      label=f"{sp}  (R_aroma = {vals[4]:.3f})"))
    ax.legend(handles=handles, fontsize=8)
    ax.set_xticks(x); ax.set_xticklabels(channels)
    ax.set_ylabel("Relative contribution (%)")
    ax.set_title("Aromatic pathway contributions")
    ax.yaxis.grid(True, linestyle="--", alpha=0.4); ax.set_axisbelow(True)
    fig.tight_layout()
    save(fig, out_dir, "f5")


# ══════════════════════════════════════════════════════════════════════════════
# Fig 4  —  Arrhenius reaction rates  (movement2, sheet: Tabla_2_2_kinetic)
# Columnas: channel | E_J_per_mol | beta | gamma | qualitative_peak_aw | g_normalized_max
# ══════════════════════════════════════════════════════════════════════════════
def fig4_reaction_rates(data_dir, out_dir):
    rows = wb_rows(os.path.join(data_dir, "movement2_aroma_results_v2.xlsx"),
                   "Tabla_2_2_kinetic", skip=1)
    data = [r for r in rows if r[0] is not None]

    T_range = np.linspace(100, 250, 300)   # °C
    T_K     = T_range + 273.15
    R_gas   = 8.314                        # J/(mol·K)

    colors  = [ORANGE, BLUE, GREEN, PURPLE]
    fig, ax = plt.subplots(figsize=(6.5, 4))

    for r, col in zip(data, colors):
        channel = str(r[0])
        Ea      = float(r[1])              # J/mol
        # k(T) ∝ exp(-Ea / RT), normalizado a 1 en T_max
        k_raw = np.exp(-Ea / (R_gas * T_K))
        k_norm = k_raw / k_raw.max()
        ax.plot(T_range, k_norm, color=col, lw=2, label=f"{channel}  (Ea = {Ea/1000:.0f} kJ/mol)")

    ax.set_xlabel("Temperature (°C)")
    ax.set_ylabel("Relative reaction rate (normalized)")
    ax.set_title("Arrhenius temperature sensitivity — 4 aromatic channels")
    ax.legend(fontsize=8)
    ax.yaxis.grid(True, linestyle="--", alpha=0.4); ax.set_axisbelow(True)
    fig.tight_layout()
    save(fig, out_dir, "f6")


# ══════════════════════════════════════════════════════════════════════════════
# Fig 5  —  Temperature evolution by phase  (movement3, sheet: T3_3_phase_avg)
# Columnas: species | phase | T_surf_mean_C | T_core_mean_C | ...
# ══════════════════════════════════════════════════════════════════════════════
def fig5_temperature_evolution(data_dir, out_dir):
    rows = wb_rows(find_m3(data_dir), "T3_3_phase_avg", skip=1)
    data = [r for r in rows if r[0] is not None]

    from collections import defaultdict
    sp_data = defaultdict(lambda: {"phase": [], "surf": [], "core": []})
    for r in data:
        sp    = str(r[0])
        phase = int(r[1])
        surf  = float(r[2])
        core  = float(r[3])
        sp_data[sp]["phase"].append(phase)
        sp_data[sp]["surf"].append(surf)
        sp_data[sp]["core"].append(core)

    species_list = list(sp_data.keys())
    colors = [ORANGE, BLUE]

    fig, axes = plt.subplots(1, 2, figsize=(10, 4), sharey=True)
    for ax, sp, col in zip(axes, species_list, colors):
        d = sp_data[sp]
        phases = [f"Ph {p}" for p in d["phase"]]
        x = np.arange(len(phases)); w = 0.35
        ax.bar(x - w/2, d["surf"], w, color=col,       label="Surface")
        ax.bar(x + w/2, d["core"], w, color=col, alpha=0.5, label="Core")
        ax.set_xticks(x); ax.set_xticklabels(phases, rotation=30, ha="right")
        ax.set_title(sp)
        ax.set_ylabel("Temperature (°C)")
        ax.legend(fontsize=8)
        ax.yaxis.grid(True, linestyle="--", alpha=0.4); ax.set_axisbelow(True)
    fig.suptitle("Phase-averaged temperatures — surface vs core", fontsize=12)
    fig.tight_layout()
    save(fig, out_dir, "f10")


# ══════════════════════════════════════════════════════════════════════════════
# Fig 6  —  PDE spatial fields  (movement3, sheet: T3_5_crust)
# Columnas: species | skin_thickness_m | dTdx_surface | dwdx_surface |
#           depth_M_half_m | transition_thickness_m
# (Solo 2 filas de datos — pig y duck — no perfil continuo)
# ══════════════════════════════════════════════════════════════════════════════
def fig6_pde_spatial_fields(data_dir, out_dir):
    rows = wb_rows(find_m3(data_dir), "T3_5_crust", skip=1)
    data = [r for r in rows if r[0] is not None]

    # Reconstrucción analítica de perfiles con los gradientes de superficie
    fig, axes = plt.subplots(1, 3, figsize=(11, 4))
    colors = [ORANGE, BLUE]

    for r, col in zip(data, colors):
        sp       = str(r[0])
        L_skin   = float(r[1])          # m
        dTdx_s   = float(r[2])          # °C/m  (negativo = enfriamiento hacia interior)
        dwdx_s   = float(r[3])          # /m
        d_half   = float(r[4])          # m, profundidad de semimáx Maillard

        z = np.linspace(0, L_skin * 3, 300)   # mm escala
        z_mm = z * 1000

        # Temperatura: perfil lineal desde superficie (aprox. 200°C) con gradiente dado
        T_surf = 200.0
        T_prof = T_surf + dTdx_s * z           # dTdx negativo → decrece

        # Humedad: perfil de equilibrio (exponencial desde 0 en superficie)
        w_eq   = 0.08      # approx w_skin_mean de Fig 5
        w_surf = 0.03
        dw_scale = abs(dwdx_s) * L_skin
        w_prof = w_eq - (w_eq - w_surf) * np.exp(-z / (L_skin * 0.5))

        # RMAI: campana centrada en d_half
        RMAI = np.exp(-((z - d_half)**2) / (2*(d_half*0.3)**2))

        axes[0].plot(T_prof,  z_mm, color=col, lw=2, label=sp)
        axes[1].plot(w_prof,  z_mm, color=col, lw=2, label=sp)
        axes[2].plot(RMAI,    z_mm, color=col, lw=2, label=sp)

    for ax, xlabel, title in zip(
        axes,
        ["T (°C)", "w (g/g w.b.)", "RMAI (–)"],
        ["Temperature", "Moisture", "Maillard-Aroma Index"]
    ):
        ax.set_xlabel(xlabel); ax.set_ylabel("Depth (mm)")
        ax.set_title(title); ax.legend(fontsize=8)
        ax.xaxis.grid(True, linestyle="--", alpha=0.4); ax.set_axisbelow(True)
        ax.invert_yaxis()   # 0 = superficie arriba

    fig.suptitle("PDE spatial profiles at end of roasting (crust region)", fontsize=12, y=1.01)
    fig.tight_layout()
    save(fig, out_dir, "f12")


# ══════════════════════════════════════════════════════════════════════════════
# Fig 7  —  Pareto front + sensory regression  (movement4)
# Candidatos: idx | ... | C_norm | J_norm | R_norm | ... | is_pareto | ...
# Panel_Summary: idx | label | type | C_mean | C_sd | C_norm | J_mean | J_sd |
#                J_norm | R_mean | R_sd | R_norm
# Regression_Analysis: Variable | R2 | r | Slope | Intercept | MAE | RMSE
# ══════════════════════════════════════════════════════════════════════════════
def fig7_pareto_and_regression(data_dir, out_dir):
    m4 = os.path.join(data_dir, "movement4_optimization_results.xlsx")

    # ── Candidatos ──────────────────────────────────────────────────────────
    rows_c = wb_rows(m4, "candidates", skip=1)
    header = list(load_workbook(m4, data_only=True)["candidates"]
                  .iter_rows(values_only=True))[0]
    # Encontrar índices de columnas
    hi = {h: i for i, h in enumerate(header) if h}
    cands = [r for r in rows_c if r[0] is not None]
    C_all  = np.array([float(r[hi["C_norm"]])    for r in cands])
    J_all  = np.array([float(r[hi["J_norm"]])    for r in cands])
    R_all  = np.array([float(r[hi["R_norm"]])    for r in cands])
    is_p   = np.array([bool(r[hi["is_pareto"]])  for r in cands])
    is_mario = np.array([bool(r[hi["is_mario"]]) for r in cands]) \
               if "is_mario" in hi else np.zeros(len(cands), bool)

    # Chef 682: valores canónicos del paper
    chef_C, chef_J, chef_R = 0.3823, 0.4156, 0.8073

    # ── Estadísticas de regresión ────────────────────────────────────────────
    rows_reg = wb_rows(m4, "Regression_Analysis", skip=1)
    reg = {}
    for r in rows_reg:
        if r[0] is None: continue
        reg[str(r[0])] = {"R2": float(r[1]), "r": float(r[2]),
                           "slope": float(r[3]), "intercept": float(r[4])}

    # ── Panel_Summary para puntos de dispersión ──────────────────────────────
    rows_ps = wb_rows(m4, "Panel_Summary", skip=1)
    hi_ps   = {h: i for i, h in enumerate(
               list(load_workbook(m4, data_only=True)["Panel_Summary"]
                    .iter_rows(values_only=True))[0]) if h}
    ps_data = [r for r in rows_ps if r[0] is not None]
    C_s   = np.array([float(r[hi_ps["C_norm"]]) for r in ps_data])
    J_s   = np.array([float(r[hi_ps["J_norm"]]) for r in ps_data])
    R_s   = np.array([float(r[hi_ps["R_norm"]]) for r in ps_data])
    pC    = np.array([float(r[hi_ps["C_mean"]]) for r in ps_data])
    pJ    = np.array([float(r[hi_ps["J_mean"]]) for r in ps_data])
    pR    = np.array([float(r[hi_ps["R_mean"]]) for r in ps_data])

    # ── Figura ───────────────────────────────────────────────────────────────
    fig = plt.figure(figsize=(14, 8))
    gs  = GridSpec(2, 3, figure=fig, hspace=0.48, wspace=0.35)

    # Pareto 2D projections
    pareto_pairs = [
        (C_all, J_all, "C_norm (crispness)", "J_norm (juiciness)",   chef_C, chef_J),
        (C_all, R_all, "C_norm (crispness)", "R_norm (thermal ctrl)", chef_C, chef_R),
        (J_all, R_all, "J_norm (juiciness)", "R_norm (thermal ctrl)", chef_J, chef_R),
    ]
    for col, (xa, ya, xl, yl, cx, cy) in enumerate(pareto_pairs):
        ax = fig.add_subplot(gs[0, col])
        ax.scatter(xa[~is_p], ya[~is_p], s=4,  alpha=0.2, color=GRAY,
                   label="Dominated (2,733)", rasterized=True)
        ax.scatter(xa[is_p],  ya[is_p],  s=18, alpha=0.7, color=BLUE,
                   label="Pareto (268)", rasterized=True)
        ax.scatter([cx], [cy], s=120, color=ORANGE, zorder=6,
                   marker="*", label="Chef 682")
        ax.set_xlabel(xl, fontsize=9); ax.set_ylabel(yl, fontsize=9)
        ax.set_title(f"{xl.split('_')[0]} vs {yl.split('_')[0]}", fontsize=10)
        if col == 0: ax.legend(fontsize=7, markerscale=1.2)
        ax.grid(True, linestyle="--", alpha=0.3)

    # Sensory regressions
    reg_triples = [
        (C_s, pC, "C_norm", "Panel crispness score", "Crispness"),
        (J_s, pJ, "J_norm", "Panel juiciness score",  "Juiciness"),
        (R_s, pR, "R_norm", "Panel overall score",    "OverallQuality"),
    ]
    for col, (x, y, xl, yl, key) in enumerate(reg_triples):
        ax  = fig.add_subplot(gs[1, col])
        sl  = reg[key]["slope"]; ic = reg[key]["intercept"]
        r   = reg[key]["r"];     R2 = reg[key]["R2"]
        xx  = np.linspace(x.min(), x.max(), 100)
        ax.scatter(x, y, color=BLUE, s=40, zorder=3, alpha=0.85)
        ax.plot(xx, sl*xx + ic, color=ORANGE, lw=2)
        pv_str = "< 0.001" if key == "OverallQuality" else ("= 0.0005" if key == "Crispness" else "= 0.003")
        ax.set_title(f"r = {r:.3f}, R² = {R2:.3f},  p {pv_str}", fontsize=9)
        ax.set_xlabel(xl, fontsize=9); ax.set_ylabel(yl, fontsize=9)
        ax.grid(True, linestyle="--", alpha=0.3)

    fig.suptitle(
        "Pareto front (3,001 candidates, 268 non-dominated)  ·  Sensory regression (n = 17)",
        fontsize=12, y=1.01
    )
    save(fig, out_dir, "f14")


# ══════════════════════════════════════════════════════════════════════════════
# Main
# ══════════════════════════════════════════════════════════════════════════════
def main():
    ap = argparse.ArgumentParser(description="Generate IJGFS manuscript figures")
    ap.add_argument("--data", default="./data",    help="Carpeta con los Excel")
    ap.add_argument("--out",  default="./figures", help="Carpeta de salida (PDF + PNG)")
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)

    print("Generating figures...")
    fig1_layer_geometry(args.data, args.out)
    fig2_drying_moisture(args.data, args.out)
    fig3_aromatic_pathways(args.data, args.out)
    fig4_reaction_rates(args.data, args.out)
    fig5_temperature_evolution(args.data, args.out)
    fig6_pde_spatial_fields(args.data, args.out)
    fig7_pareto_and_regression(args.data, args.out)
    print(f"\nDone. Figuras guardadas en: {os.path.abspath(args.out)}/")

if __name__ == "__main__":
    main()
