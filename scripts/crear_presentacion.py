from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from sklearn.compose import ColumnTransformer
from sklearn.dummy import DummyClassifier
from sklearn.ensemble import RandomForestClassifier
from sklearn.inspection import permutation_importance
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (
    accuracy_score,
    balanced_accuracy_score,
    confusion_matrix,
    f1_score,
)
from sklearn.model_selection import train_test_split
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt
from scipy.stats import chi2_contingency, spearmanr


ROOT = Path(__file__).resolve().parents[1]
DATA_PATH = ROOT / "data" / "online food delivery dataset.csv"
OUTPUT_PATH = ROOT / "presentacion_final_delivery.pptx"

COLORS = {
    "ink": "172A2A",
    "muted": "536665",
    "paper": "F6F4EE",
    "white": "FFFFFF",
    "teal": "087E78",
    "teal_light": "D9ECE7",
    "coral": "D45D43",
    "coral_light": "F5E1D9",
    "gold": "D9A441",
    "line": "D8DED8",
    "blue": "326B83",
}
FONT = "Aptos"


def rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def load_data() -> pd.DataFrame:
    data = pd.read_csv(DATA_PATH)
    data = data.loc[:, ~data.columns.astype(str).str.startswith("Unnamed")].copy()
    for column in data.select_dtypes(include=["object", "str"]).columns:
        data[column] = data[column].str.strip()
    return data


def cramers_v(table: pd.DataFrame) -> float:
    chi2 = chi2_contingency(table, correction=False)[0]
    n = table.to_numpy().sum()
    rows, columns = table.shape
    phi2 = chi2 / n
    phi2_corrected = max(0.0, phi2 - ((columns - 1) * (rows - 1)) / (n - 1))
    rows_corrected = rows - ((rows - 1) ** 2) / (n - 1)
    columns_corrected = columns - ((columns - 1) ** 2) / (n - 1)
    denominator = max(1.0, min(columns_corrected - 1, rows_corrected - 1))
    return float(np.sqrt(phi2_corrected / denominator))


def duplicate_summary(data: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    duplicate_rows = data.loc[data.duplicated(keep=False)].copy()
    duplicate_rows["CSV row"] = duplicate_rows.index + 2
    groups = (
        duplicate_rows.groupby(list(data.columns), dropna=False, sort=False)
        .agg(count=("CSV row", "size"), rows=("CSV row", list))
        .reset_index()
        .sort_values("count", ascending=False, kind="stable")
    )
    frequency = groups.groupby("count").size().sort_index(ascending=False)
    return groups, frequency


def add_text(slide, text: str, x: float, y: float, w: float, h: float,
             size: float = 16, color: str | None = None, bold: bool = False,
             align=PP_ALIGN.LEFT, font: str = FONT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = Inches(0.02)
    frame.margin_right = Inches(0.02)
    frame.margin_top = Inches(0.01)
    frame.margin_bottom = Inches(0.01)
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color or COLORS["ink"])
    return box


def add_rich_lines(slide, lines: list[tuple[str, str]], x: float, y: float,
                   w: float, h: float, size: float = 16):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = Inches(0.02)
    frame.margin_right = Inches(0.03)
    for index, (lead, body) in enumerate(lines):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.space_after = Pt(12)
        first = paragraph.add_run()
        first.text = lead
        first.font.name = FONT
        first.font.size = Pt(size)
        first.font.bold = True
        first.font.color.rgb = rgb(COLORS["teal"])
        second = paragraph.add_run()
        second.text = body
        second.font.name = FONT
        second.font.size = Pt(size)
        second.font.color.rgb = rgb(COLORS["ink"])
    return box


def add_base(prs: Presentation, title: str, kicker: str = "HALLAZGOS"):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(COLORS["paper"])
    add_text(slide, kicker.upper(), 0.62, 0.34, 5.8, 0.28, 10,
             COLORS["teal"], True)
    add_text(slide, title, 0.62, 0.68, 12.0, 0.62, 27,
             COLORS["ink"], True)
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.62), Inches(1.42), Inches(12.05), Inches(0.025)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = rgb(COLORS["line"])
    line.line.fill.background()
    add_text(slide, "Online food delivery | Análisis exploratorio", 0.62, 7.13,
             8.7, 0.2, 9, COLORS["muted"])
    add_text(slide, f"{len(prs.slides):02d}", 12.05, 7.1, 0.55, 0.22,
             9, COLORS["muted"], align=PP_ALIGN.RIGHT)
    return slide


def add_metric(slide, value: str, label: str, x: float, y: float,
               accent: str = "teal", width: float = 2.75):
    color = COLORS[accent]
    add_text(slide, value, x, y, width, 0.68, 32, color, True)
    add_text(slide, label, x, y + 0.68, width, 0.64, 13, COLORS["muted"])


def chart_to_png(fig, folder: Path, name: str) -> Path:
    path = folder / f"{name}.png"
    fig.savefig(path, dpi=190, bbox_inches="tight", facecolor=f"#{COLORS['paper']}")
    plt.close(fig)
    return path


def create_charts(data: pd.DataFrame, folder: Path) -> dict[str, Path]:
    sns.set_theme(style="whitegrid", font="DejaVu Sans")
    charts: dict[str, Path] = {}
    analysis_data = data.drop_duplicates().reset_index(drop=True)

    associations = {
        "Ocupacion e ingreso": pd.crosstab(analysis_data["Occupation"], analysis_data["Monthly Income"]),
        "Educacion y ocupacion": pd.crosstab(
            analysis_data["Educational Qualifications"], analysis_data["Occupation"]
        ),
        "Estado civil e ingreso": pd.crosstab(analysis_data["Marital Status"], analysis_data["Monthly Income"]),
        "Resultado y feedback": pd.crosstab(analysis_data["Output"], analysis_data["Feedback"]),
    }
    values = {name: cramers_v(table) for name, table in associations.items()}
    ordered = sorted(values.items(), key=lambda item: item[1])
    fig, ax = plt.subplots(figsize=(8.8, 4.3))
    names = [item[0] for item in ordered]
    scores = [item[1] for item in ordered]
    bars = ax.barh(names, scores, color=[f"#{COLORS[key]}" for key in ("blue", "gold", "coral", "teal")])
    ax.set_xlim(0, 0.68)
    ax.set_xlabel("V de Cramer corregida")
    ax.set_ylabel("")
    ax.set_title("Fuerza de asociacion entre pares categoricos", loc="left", weight="bold")
    ax.grid(axis="x", alpha=0.2)
    ax.grid(axis="y", visible=False)
    for bar, score in zip(bars, scores):
        ax.text(score + 0.012, bar.get_y() + bar.get_height() / 2,
                f"{score:.3f}", va="center", fontsize=10, color=f"#{COLORS['ink']}")
    charts["cramer"] = chart_to_png(fig, folder, "cramer")

    income_order = ["No Income", "Below Rs.10000", "10001 to 25000", "25001 to 50000", "More than 50000"]
    occupation_income = pd.crosstab(analysis_data["Occupation"], analysis_data["Monthly Income"], normalize="index") * 100
    occupation_income = occupation_income.reindex(columns=income_order, fill_value=0)
    fig, ax = plt.subplots(figsize=(8.5, 3.8))
    occupation_income.plot(
        kind="barh", stacked=True, ax=ax,
        color=[f"#{COLORS[key]}" for key in ("ink", "blue", "teal", "gold", "coral")],
        width=0.67,
    )
    ax.set_xlim(0, 100)
    ax.set_xlabel("Porcentaje dentro de cada ocupacion")
    ax.set_ylabel("")
    ax.set_title("Ingreso mensual segun ocupacion", loc="left", weight="bold")
    ax.legend(title="Tramo de ingreso", bbox_to_anchor=(1.02, 1), loc="upper left", fontsize=8)
    ax.grid(axis="x", alpha=0.2)
    ax.grid(axis="y", visible=False)
    charts["income"] = chart_to_png(fig, folder, "income")

    order = ["New", "Regular", "Frequent"]
    fig, ax = plt.subplots(figsize=(7.3, 3.8))
    sns.stripplot(
        data=analysis_data, x="Customer Type", y="Family size", order=order,
        jitter=0.15, alpha=0.42, color=f"#{COLORS['teal']}", ax=ax,
    )
    sns.pointplot(
        data=analysis_data, x="Customer Type", y="Family size", order=order,
        errorbar=("ci", 95), color=f"#{COLORS['coral']}", markers="D",
        linestyles="--", ax=ax,
    )
    ax.set_title("Tamano familiar por tipo de cliente", loc="left", weight="bold")
    ax.set_xlabel("Tipo de cliente (orden ordinal)")
    ax.set_ylabel("Tamano familiar")
    ax.grid(axis="y", alpha=0.2)
    ax.grid(axis="x", visible=False)
    charts["family"] = chart_to_png(fig, folder, "family")

    output_feedback = pd.crosstab(analysis_data["Output"], analysis_data["Feedback"])
    _, _, _, expected = chi2_contingency(output_feedback, correction=False)
    residuals = (output_feedback - expected) / np.sqrt(expected)
    limit = max(float(np.abs(residuals.to_numpy()).max()), 1.0)
    fig, ax = plt.subplots(figsize=(5.3, 3.3))
    sns.heatmap(
        residuals, annot=True, fmt=".2f", cmap="RdBu_r", center=0,
        vmin=-limit, vmax=limit, linewidths=0.5,
        cbar_kws={"label": "Residuo de Pearson"}, ax=ax,
    )
    ax.set_title("Resultado del pedido y feedback", loc="left", weight="bold")
    ax.set_xlabel("Feedback")
    ax.set_ylabel("Output")
    charts["feedback"] = chart_to_png(fig, folder, "feedback")

    duplicate_rows = data.loc[data.duplicated(keep=False)]
    groups = duplicate_rows.groupby(list(data.columns), dropna=False, sort=False).size()
    freq = groups.value_counts().sort_index(ascending=False)
    fig, ax = plt.subplots(figsize=(6.6, 3.7))
    categories = [f"{count} veces" for count in freq.index]
    bars = ax.bar(categories, freq.values, color=[f"#{COLORS[key]}" for key in ("coral", "gold", "teal", "blue")])
    ax.set_ylabel("Numero de grupos de registros")
    ax.set_title("Frecuencia de los grupos duplicados", loc="left", weight="bold")
    ax.grid(axis="y", alpha=0.2)
    ax.grid(axis="x", visible=False)
    for bar, value in zip(bars, freq.values):
        ax.text(bar.get_x() + bar.get_width() / 2, value + 0.5,
                str(value), ha="center", va="bottom", weight="bold")
    charts["duplicates"] = chart_to_png(fig, folder, "duplicates")
    return charts


def create_model_analysis(data: pd.DataFrame, folder: Path) -> tuple[dict[str, Path], dict]:
    model_data = data.drop_duplicates().reset_index(drop=True)
    target = "Output"
    excluded_columns = [target, "Feedback", "Customer Type", "Pin code"]
    X = model_data.drop(columns=excluded_columns)
    y = model_data[target]

    X_train_val, X_test, y_train_val, y_test = train_test_split(
        X, y, test_size=0.2, random_state=42, stratify=y
    )
    X_train, X_val, y_train, y_val = train_test_split(
        X_train_val, y_train_val, test_size=0.25, random_state=42, stratify=y_train_val
    )

    categorical_columns = X_train.select_dtypes(
        include=["object", "category", "string"]
    ).columns.tolist()
    numeric_columns = [column for column in X_train.columns if column not in categorical_columns]
    preprocessor = ColumnTransformer(
        transformers=[
            ("categorical", OneHotEncoder(handle_unknown="ignore"), categorical_columns),
            ("numeric", StandardScaler(), numeric_columns),
        ]
    )
    model_factories = {
        "Baseline (clase frecuente)": lambda: DummyClassifier(strategy="most_frequent"),
        "Regresion logistica": lambda: LogisticRegression(
            class_weight="balanced", max_iter=2000, random_state=42
        ),
        "Random forest": lambda: RandomForestClassifier(
            class_weight="balanced", n_estimators=300, random_state=42
        ),
    }

    fitted_models = {}
    metric_rows = []
    for model_name, make_model in model_factories.items():
        estimator = make_model()
        model = (
            estimator
            if model_name == "Baseline (clase frecuente)"
            else Pipeline(steps=[("preprocessor", preprocessor), ("classifier", estimator)])
        )
        model.fit(X_train, y_train)
        fitted_models[model_name] = model
        row = {"Modelo": model_name}
        for split_name, features, labels in [
            ("Train", X_train, y_train),
            ("Validacion", X_val, y_val),
        ]:
            predictions = model.predict(features)
            row[f"{split_name} accuracy"] = accuracy_score(labels, predictions)
            row[f"{split_name} balanced accuracy"] = balanced_accuracy_score(labels, predictions)
            row[f"{split_name} F1 No"] = f1_score(
                labels, predictions, pos_label="No", zero_division=0
            )
        metric_rows.append(row)

    results = pd.DataFrame(metric_rows).set_index("Modelo")
    selected_name = results.sort_values(
        ["Validacion balanced accuracy", "Validacion F1 No"], ascending=False
    ).index[0]
    selected_model = fitted_models[selected_name]

    charts: dict[str, Path] = {}
    counts = model_data[target].value_counts().reindex(["Yes", "No"])
    percentages = counts / counts.sum() * 100
    fig, ax = plt.subplots(figsize=(6.1, 3.8))
    bars = ax.bar(["Yes", "No"], counts.values,
                  color=[f"#{COLORS['teal']}", f"#{COLORS['coral']}"], width=0.58)
    for bar, count, percentage in zip(bars, counts.values, percentages.values):
        ax.text(bar.get_x() + bar.get_width() / 2, count + 3,
                f"{count} ({percentage:.1f}%)", ha="center", weight="bold")
    ax.set_ylim(0, counts.max() * 1.2)
    ax.set_ylabel("Clientes unicos")
    ax.set_title("Distribucion de la variable objetivo", loc="left", weight="bold")
    ax.grid(axis="y", alpha=0.2)
    ax.set_axisbelow(True)
    charts["target"] = chart_to_png(fig, folder, "target")

    fig, axes = plt.subplots(1, 2, figsize=(10, 4.2))
    for ax, metric, title in zip(
        axes,
        ["balanced accuracy", "F1 No"],
        ["Balanced accuracy", "F1 para la clase minoritaria No"],
    ):
        results[[f"Train {metric}", f"Validacion {metric}"]].plot(
            kind="bar", ax=ax, color=[f"#{COLORS['blue']}", f"#{COLORS['coral']}"], width=0.75
        )
        ax.set_title(title, fontsize=11, weight="bold")
        ax.set_xlabel("")
        ax.set_ylabel("Puntaje")
        ax.set_ylim(0, 1.08)
        ax.tick_params(axis="x", labelrotation=25, labelsize=8)
        ax.legend(["Train", "Validacion"], frameon=False, fontsize=8)
        ax.grid(axis="y", alpha=0.2)
        ax.set_axisbelow(True)
    fig.suptitle("Desempeno en entrenamiento y validacion", weight="bold")
    fig.tight_layout()
    charts["validation"] = chart_to_png(fig, folder, "validation")

    importance = permutation_importance(
        selected_model,
        X_val,
        y_val,
        scoring="balanced_accuracy",
        n_repeats=30,
        random_state=42,
        n_jobs=-1,
    )
    importance_df = pd.DataFrame(
        {
            "Variable": X_val.columns,
            "Importancia media": importance.importances_mean,
            "Desviacion": importance.importances_std,
        }
    ).sort_values("Importancia media", ascending=False)
    top_importance = importance_df.head(9).sort_values("Importancia media")
    fig, ax = plt.subplots(figsize=(7.1, 4.4))
    ax.barh(
        top_importance["Variable"],
        top_importance["Importancia media"],
        xerr=top_importance["Desviacion"],
        color=f"#{COLORS['teal']}",
    )
    ax.axvline(0, color=f"#{COLORS['ink']}", linewidth=1)
    ax.set_xlabel("Cambio en balanced accuracy al permutar")
    ax.set_title("Importancia por permutacion en validacion", loc="left", weight="bold")
    ax.grid(axis="x", alpha=0.2)
    ax.set_axisbelow(True)
    charts["importance"] = chart_to_png(fig, folder, "importance")

    from sklearn.base import clone

    final_model = clone(selected_model)
    final_model.fit(X_train_val, y_train_val)
    test_predictions = final_model.predict(X_test)
    test_metrics = {
        "Accuracy": accuracy_score(y_test, test_predictions),
        "Balanced accuracy": balanced_accuracy_score(y_test, test_predictions),
        "F1 clase No": f1_score(y_test, test_predictions, pos_label="No", zero_division=0),
    }
    matrix = confusion_matrix(y_test, test_predictions, labels=["No", "Yes"])
    fig, ax = plt.subplots(figsize=(5.0, 3.8))
    sns.heatmap(
        matrix,
        annot=True,
        fmt="d",
        cmap="YlGnBu",
        xticklabels=["No", "Yes"],
        yticklabels=["No", "Yes"],
        ax=ax,
    )
    ax.set_title("Matriz de confusion: test final", loc="left", weight="bold")
    ax.set_xlabel("Prediccion")
    ax.set_ylabel("Valor real")
    charts["confusion"] = chart_to_png(fig, folder, "confusion")

    analysis = {
        "rows": len(model_data),
        "features": X.columns.tolist(),
        "split_summary": {
            "train": len(X_train), "validation": len(X_val), "test": len(X_test),
            "test_no": int((y_test == "No").sum()),
            "test_yes": int((y_test == "Yes").sum()),
        },
        "results": results,
        "selected_name": selected_name,
        "balanced_accuracy_gap": float(
            results.loc[selected_name, "Train balanced accuracy"]
            - results.loc[selected_name, "Validacion balanced accuracy"]
        ),
        "importance": importance_df,
        "test_metrics": test_metrics,
        "confusion": matrix,
    }
    return charts, analysis


def build_presentation(data: pd.DataFrame, charts: dict[str, Path], model: dict) -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = "Database Analysis"
    prs.core_properties.subject = "Customer patterns and data quality"
    prs.core_properties.author = "Regina Hernández y Matias Cohen"

    # Slide 1: title.
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(COLORS["ink"])
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.2), Inches(7.5)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = rgb(COLORS["coral"])
    accent.line.fill.background()
    add_text(slide, "DATABASE ANALYSIS", 0.85, 0.75, 5.8, 0.3,
             11, "8BC7B8", True)
    add_text(slide, "Lo que dicen\nlos clientes", 0.82, 1.45,
             8.8, 2.25, 30, COLORS["white"], True)
    add_text(slide, "Regina Hernández y Matias Cohen", 0.87,
             4.05, 8.5, 0.55, 19, "D6E2DC")
    add_text(slide, f"{len(data)} respuestas  |  13 variables  |  Bengaluru, India",
             0.87, 6.47, 8.6, 0.35, 13, "B7C9C2")
    circle = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(10.1), Inches(1.45), Inches(1.75), Inches(1.75)
    )

    # Slide 2: why this dataset.
    slide = add_base(prs, "¿Por qué lo elegimos?", "Motivo")
    add_metric(slide, "388", "respuestas originales", 0.8, 1.9)
    add_metric(slide, "13", "variables clave", 3.6, 1.9, "coral")
    add_metric(slide, "4", "patrones claros", 6.4, 1.9, "blue")
    add_rich_lines(slide, [
        ("Es útil porque: ", "muestra decisiones reales de clientes en una app de comida."),
        ("Se entiende fácilmente: ", "hay perfiles, hábitos y experiencias que son muy claros de ver."),
        ("Sirve para decidir: ", "los resultados ayudan a explicar el comportamiento y no quedan solo en números."),
    ], 0.9, 3.8, 11.2, 2.7, 17)
    add_text(slide, "Elegimos este dataset porque tiene datos cotidianos, claros y fáciles de interpretar.",
             0.8, 6.55, 11.6, 0.28, 11, COLORS["muted"])

    circle.fill.solid()
    circle.fill.fore_color.rgb = rgb(COLORS["teal"])
    circle.line.fill.background()
    add_text(slide, "388", 10.1, 1.9, 1.75, 0.55, 27, COLORS["white"], True,
             PP_ALIGN.CENTER)
    add_text(slide, "respuestas", 10.15, 2.48, 1.65, 0.28, 11, COLORS["white"],
             align=PP_ALIGN.CENTER)

    # Slide 2: scope and method.
    slide = add_base(prs, "Qué revisamos", "Alcance")
    add_metric(slide, "388", "respuestas del dataset", 0.75, 1.9)
    add_metric(slide, "13", "variables analizadas", 3.8, 1.9, "coral")
    add_metric(slide, "5", "relaciones exploradas", 6.85, 1.9, "blue")
    add_metric(slide, "66", "grupos duplicados", 9.9, 1.9, "gold")
    add_rich_lines(slide, [
        ("Gráfico de barras: ", "muestra qué tan fuerte se relacionan las categorías."),
        ("Gráfico apilado: ", "compara los ingresos según el tipo de trabajo."),
        ("Mapa de calor: ", "ayuda a ver dónde los datos se salen de lo normal."),
    ], 0.82, 3.42, 11.4, 2.35, 15)

    # Slide 3: variable types.
    slide = add_base(prs, "Tipos de variables", "Variables")
    headers = ["Variable", "Tipo de dato", "Ejemplo o categoría"]
    rows = [
        ["Nivel educativo", "Categórica ordinal", "Uneducated, School, Graduate, Post Graduate, Ph.D."],
        ["Ocupación", "Categórica nominal", "Student, Employee, Self Employeed, House wife"],
        ["Ingreso mensual", "Categórica ordinal", "No Income, Below Rs.10000, 10k–25k, 25k–50k, 50k+"],
        ["Estado civil", "Categórica nominal", "Single, Married, Prefer not to say"],
        ["Tipo de cliente", "Categórica ordinal", "New, Regular, Frequent"],
        ["Tamaño familiar", "Cuantitativa discreta", "1 a 6 personas"],
        ["Resultado / feedback", "Categóricas nominales", "Sí/No y positivo/negativo"],
    ]
    table = slide.shapes.add_table(
        len(rows) + 1, 3, Inches(0.75), Inches(1.78), Inches(11.8), Inches(4.75)
    ).table
    table.columns[0].width = Inches(2.25)
    table.columns[1].width = Inches(3.05)
    table.columns[2].width = Inches(6.5)
    for col, header in enumerate(headers):
        cell = table.cell(0, col)
        cell.text = header
    for row_index, row in enumerate(rows, start=1):
        for col_index, value in enumerate(row):
            table.cell(row_index, col_index).text = value
    for row_index in range(len(rows) + 1):
        for col_index in range(3):
            cell = table.cell(row_index, col_index)
            cell.margin_left = Inches(0.13)
            cell.margin_right = Inches(0.1)
            cell.margin_top = Inches(0.04)
            cell.margin_bottom = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(COLORS["ink"] if row_index == 0 else (COLORS["white"] if row_index % 2 else "EEF1EB"))
            for paragraph in cell.text_frame.paragraphs:
                paragraph.space_after = Pt(0)
                for run in paragraph.runs:
                    run.font.name = FONT
                    run.font.size = Pt(12 if row_index else 12)
                    run.font.bold = row_index == 0
                    run.font.color.rgb = rgb(COLORS["white"] if row_index == 0 else COLORS["ink"])
    add_text(slide, "Datos: categóricos, ordinales y discretos. Eso cambia la forma de analizarlos.",
             0.82, 6.62, 11.5, 0.27, 11, COLORS["muted"])

    # Slide 4: association ranking.
    slide = add_base(prs, "Las relaciones más claras", "Patrones")
    slide.shapes.add_picture(str(charts["cramer"]), Inches(0.65), Inches(1.65), width=Inches(7.6))
    add_rich_lines(slide, [
        ("Gráfico de barras: ", "muestra qué tan fuerte se relacionan las categorías."),
        ("Trabajo e ingreso: ", "es la relación más clara y se ve muy marcada."),
        ("Feedback y resultado: ", "también van juntos y ayudan a entender la experiencia del cliente."),
    ], 8.35, 1.9, 4.15, 4.4, 14)
    add_text(slide, "Es una comparación visual, no una prueba de causa y efecto. | Tipo de gráfica: barras. Datos: categóricos.",
             0.83, 6.52, 11.7, 0.3, 11, COLORS["muted"])

    # Slide 5: occupation and income.
    slide = add_base(prs, "Ingreso según trabajo", "Insight 1")
    slide.shapes.add_picture(str(charts["income"]), Inches(0.6), Inches(1.72), width=Inches(8.25))
    add_rich_lines(slide, [
        ("Gráfico apilado: ", "compara los rangos de ingreso por tipo de trabajo."),
        ("Estudiantes: ", "la mayoría dice no tener ingresos."),
        ("Empleados: ", "muchos están en ingresos medios y altos."),
    ], 9.0, 1.9, 3.55, 4.55, 13)
    add_text(slide, "Los ingresos están ordenados por rangos, así que se comparan por grupos, no como valores exactos. | Tipo de gráfica: apilado. Datos: categóricos y ordinales.",
             0.83, 6.56, 10.7, 0.27, 11, COLORS["muted"])

    # Slide 6: customer type and family size.
    slide = add_base(prs, "Tamaño familiar según tipo de cliente", "Insight 2")
    slide.shapes.add_picture(str(charts["family"]), Inches(0.65), Inches(1.75), width=Inches(7.5))
    family_customer_v = cramers_v(pd.crosstab(data["Family size"], data["Customer Type"]))
    add_metric(slide, f"{family_customer_v:.3f}", "V de Cramer corregida", 8.55, 1.95, "coral", 3.0)
    add_rich_lines(slide, [
        ("Gráfico de dispersión: ", "muestra la tendencia general de cada grupo."),
        ("Patrón en estos datos: ", "tamaño familiar 1 = New, 2–3 = Regular y 4–6 = Frequent; podría ser una regla de codificación del dataset."),
    ], 8.55, 3.42, 3.8, 2.65, 14)
    add_text(slide, "Es una comparación sencilla entre grupos ordenados y tamaños familiares. | Tipo de gráfica: dispersión. Datos: categórico ordinal + numérico discreto.",
             0.83, 6.54, 11.7, 0.27, 11, COLORS["muted"])

    # Slide 7: output and feedback.
    slide = add_base(prs, "Feedback y resultado", "Insight 3")
    slide.shapes.add_picture(str(charts["feedback"]), Inches(0.8), Inches(1.9), width=Inches(5.1))
    unique_data = data.drop_duplicates().reset_index(drop=True)
    feedback_by_output = pd.crosstab(unique_data["Output"], unique_data["Feedback"], normalize="index") * 100
    positive_yes = feedback_by_output.loc["Yes", "Positive"]
    negative_no = feedback_by_output.loc["No", "Negative"]
    add_metric(slide, f"{positive_yes:.0f}%", "de resultados Yes tuvieron feedback positivo", 6.45, 1.9, "teal", 5.5)
    add_metric(slide, f"{negative_no:.0f}%", "de resultados No tuvieron feedback negativo", 6.45, 3.0, "coral", 5.5)
    add_text(slide, "Mapa de calor: las celdas más claras y más oscuras muestran dónde los datos están por encima o por debajo de lo esperado.",
             6.48, 4.45, 5.3, 1.1, 14, COLORS["ink"])
    add_text(slide, "Esto muestra una relación importante entre la experiencia del cliente y el resultado final. | Tipo de gráfica: mapa de calor. Datos: categóricos.",
             0.83, 6.54, 11.7, 0.27, 11, COLORS["muted"])

    # Slide 8: duplicates.
    duplicate_groups, frequency = duplicate_summary(data)
    duplicated_count = int(duplicate_groups["count"].sum())
    extra_copies = int((duplicate_groups["count"] - 1).sum())
    slide = add_base(prs, "Registros duplicados", "Calidad de datos")
    slide.shapes.add_picture(str(charts["duplicates"]), Inches(0.8), Inches(1.8), width=Inches(6.25))
    add_metric(slide, f"{len(duplicate_groups)}", "grupos duplicados", 7.45, 1.9, "coral", 2.45)
    add_metric(slide, f"{duplicated_count}", "filas dentro de esos grupos", 9.95, 1.9, "teal", 2.5)
    add_metric(slide, f"{extra_copies}", "copias extra", 7.45, 3.3, "blue", 5.0)
    frequency_text = "  |  ".join(f"{int(count)} veces: {int(groups)} grupos" for count, groups in frequency.items())
    add_text(slide, frequency_text, 0.85, 5.68, 11.6, 0.55, 14, COLORS["ink"], True)
    add_text(slide, "Este análisis de duplicados es exacto: compara filas completas y elimina la columna vacía sin nombre. | Tipo de gráfica: barras de frecuencia. Datos: conteos discretos.",
             0.85, 6.42, 11.6, 0.38, 10, COLORS["muted"])

    # Slide 9: conclusions.
    slide = add_base(prs, "Lo que aprendimos", "Conclusión")
    add_rich_lines(slide, [
        ("Perfiles: ", "el trabajo y el ingreso ayudan a explicar diferentes tipos de clientes de forma clara."),
        ("Patrones: ", "el tamaño de la familia y el tipo de cliente siguen una tendencia útil, y el feedback también coincide con el resultado."),
        ("Importante: ", "hay filas duplicadas en el dataset, así que antes de decidir hay que limpiarlo."),
        ("Siguiente paso: ", "volver a analizar después de limpiar los duplicados y comparar los resultados."),
    ], 0.9, 1.92, 11.4, 3.55, 17)
    add_text(slide, "Este proyecto se basa en el análisis de notebooks y en la revisión de registros duplicados.",
             0.9, 6.32, 11.3, 0.38, 11, COLORS["muted"])

    # Slide 10: prediction task and target balance.
    slide = add_base(prs, "¿Podemos anticipar una recompra?", "Modelo")
    slide.shapes.add_picture(str(charts["target"]), Inches(0.65), Inches(1.72), width=Inches(6.2))
    add_metric(slide, str(model["rows"]), "registros unicos", 7.25, 1.88, "teal", 2.3)
    add_metric(slide, str(len(model["features"])), "variables predictoras", 10.0, 1.88, "coral", 2.65)
    add_rich_lines(slide, [
        ("Objetivo y: ", "Output, es decir, si la persona volvio a ordenar (Yes/No)."),
        ("Problema: ", "clasificacion binaria; Yes representa 76.1% y No 23.9%."),
        ("Predictores X: ", "perfil del cliente y ubicacion, disponibles antes de conocer el resultado."),
        ("Exclusiones: ", "Feedback podria ser posterior; Customer Type replica Family size; Pin code es identificador."),
    ], 7.25, 3.55, 5.25, 2.65, 13)
    split = model["split_summary"]
    add_text(slide, f"Split estratificado: {split['train']} train / {split['validation']} validacion / {split['test']} test.",
             0.85, 6.22, 11.3, 0.3, 11, COLORS["ink"], True)
    add_text(slide, "La distribucion desigual hace necesario comparar mas que accuracy.",
             0.85, 6.55, 11.3, 0.3, 11, COLORS["muted"])

    # Slide 11: train and validation model comparison.
    slide = add_base(prs, "Comparamos los modelos sin tocar el test", "Validacion")
    slide.shapes.add_picture(str(charts["validation"]), Inches(0.55), Inches(1.75), width=Inches(6.55))
    results = model["results"]
    headers = ["Modelo", "Train BA", "Val BA", "Val F1 No"]
    model_names = ["Baseline (clase frecuente)", "Regresion logistica", "Random forest"]
    rows = [[
        name,
        f"{results.loc[name, 'Train balanced accuracy']:.3f}",
        f"{results.loc[name, 'Validacion balanced accuracy']:.3f}",
        f"{results.loc[name, 'Validacion F1 No']:.3f}",
    ] for name in model_names]
    table = slide.shapes.add_table(
        len(rows) + 1, len(headers), Inches(7.25), Inches(1.95), Inches(5.25), Inches(2.85)
    ).table
    for column, width in enumerate([2.35, 0.95, 0.95, 1.0]):
        table.columns[column].width = Inches(width)
    for row_index, values in enumerate([headers] + rows):
        for column_index, value in enumerate(values):
            cell = table.cell(row_index, column_index)
            cell.text = value
            cell.margin_left = Inches(0.06)
            cell.margin_right = Inches(0.04)
            cell.margin_top = Inches(0.03)
            cell.margin_bottom = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(
                COLORS["ink"] if row_index == 0 else
                (COLORS["teal_light"] if values[0] == model["selected_name"] else COLORS["white"])
            )
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.name = FONT
                    run.font.size = Pt(10)
                    run.font.bold = row_index == 0 or values[0] == model["selected_name"]
                    run.font.color.rgb = rgb(COLORS["white"] if row_index == 0 else COLORS["ink"])
    add_rich_lines(slide, [
        ("Seleccion: ", f"{model['selected_name']} por balanced accuracy de validacion."),
        ("Cautela: ", f"la brecha train-validacion es {model['balanced_accuracy_gap']:.3f}; hay indicios de sobreajuste."),
    ], 7.3, 5.15, 5.1, 1.1, 12)
    add_text(slide, "BA = balanced accuracy; F1 No resume precision y recall de la clase minoritaria.",
             0.82, 6.58, 11.8, 0.27, 10, COLORS["muted"])

    # Slide 12: final test results.
    slide = add_base(prs, "Resultado final en datos no vistos", "Test")
    slide.shapes.add_picture(str(charts["confusion"]), Inches(0.8), Inches(1.72), width=Inches(5.35))
    test_metrics = model["test_metrics"]
    add_metric(slide, f"{test_metrics['Accuracy']:.3f}", "accuracy", 6.55, 1.95, "blue", 2.2)
    add_metric(slide, f"{test_metrics['Balanced accuracy']:.3f}", "balanced accuracy", 9.45, 1.95, "teal", 2.65)
    add_metric(slide, f"{test_metrics['F1 clase No']:.3f}", "F1 clase No", 6.55, 3.3, "coral", 2.5)
    matrix = model["confusion"]
    add_rich_lines(slide, [
        ("Lectura: ", f"detectamos {matrix[0, 0]} de {matrix[0].sum()} casos No y {matrix[1, 1]} de {matrix[1].sum()} casos Yes."),
        ("Tamaño del test: ", f"{model['split_summary']['test']} personas, incluidas {model['split_summary']['test_no']} No."),
        ("Interpretacion: ", "un acierto o error cambia mucho la metrica de No con una muestra tan pequeña."),
    ], 6.55, 4.45, 5.85, 1.8, 13)
    add_text(slide, "El test se utilizo una sola vez, despues de elegir el modelo con validacion.",
             0.82, 6.58, 11.8, 0.27, 10, COLORS["muted"])

    # Slide 13: feature importance.
    slide = add_base(prs, "¿Qué variables aportaron señal?", "Importancia")
    slide.shapes.add_picture(str(charts["importance"]), Inches(0.58), Inches(1.68), width=Inches(7.25))
    top_features = ", ".join(model["importance"].head(3)["Variable"].tolist())
    add_rich_lines(slide, [
        ("Metodo: ", "permutacion en validacion; mide cuanto cae balanced accuracy al mezclar cada variable."),
        ("Mayores señales: ", f"{top_features}."),
        ("Relacion con EDA: ", "ocupacion e ingreso ya mostraban asociacion; el modelo tambien utiliza edad y ubicacion."),
        ("Limite: ", "las barras de error son amplias; no es evidencia causal ni un ranking estable."),
    ], 8.0, 1.9, 4.45, 4.15, 13)
    add_text(slide, "Importancia calculada sobre validacion, no sobre el conjunto de test.",
             0.82, 6.58, 11.8, 0.27, 10, COLORS["muted"])

    # Slide 14: conclusion.
    slide = add_base(prs, "Conclusión: señal útil, evidencia todavía limitada", "Cierre")
    clean_data = data.drop_duplicates().reset_index(drop=True)
    occupation_income_v = cramers_v(pd.crosstab(clean_data["Occupation"], clean_data["Monthly Income"]))
    education_occupation_v = cramers_v(pd.crosstab(
        clean_data["Educational Qualifications"], clean_data["Occupation"]
    ))
    add_rich_lines(slide, [
        ("EDA: ", f"ocupacion-ingreso (V corregida = {occupation_income_v:.3f}) y educacion-ocupacion (V = {education_occupation_v:.3f}) aportan contexto."),
        ("Modelo: ", f"{model['selected_name']} supera al baseline en balanced accuracy, pero su brecha de {model['balanced_accuracy_gap']:.3f} advierte sobreajuste."),
        ("Resultado: ", f"en test, balanced accuracy = {test_metrics['Balanced accuracy']:.3f} y F1 para No = {test_metrics['F1 clase No']:.3f}; es exploratorio, no listo para decisiones reales."),
        ("Siguiente paso: ", "reunir mas respuestas independientes, validar con particiones estratificadas repetidas y revisar el momento real de cada variable."),
    ], 0.9, 1.95, 11.4, 3.95, 16)
    add_text(slide, "Asociacion y capacidad predictiva no demuestran causalidad.",
             0.9, 6.43, 11.3, 0.32, 11, COLORS["muted"])

    prs.save(OUTPUT_PATH)


def main() -> None:
    data = load_data()
    with TemporaryDirectory(prefix="delivery_presentation_") as temp_dir:
        charts = create_charts(data, Path(temp_dir))
        model_charts, model_analysis = create_model_analysis(data, Path(temp_dir))
        charts.update(model_charts)
        build_presentation(data, charts, model_analysis)
    print(f"Presentacion creada: {OUTPUT_PATH}")
    print(f"Registros: {len(data)} | Variables: {len(data.columns)}")
    duplicate_groups, _ = duplicate_summary(data)
    print(
        f"Grupos duplicados: {len(duplicate_groups)} | "
        f"Filas implicadas: {int(duplicate_groups['count'].sum())} | "
        f"Copias extra: {int((duplicate_groups['count'] - 1).sum())}"
    )
    print(f"Modelo seleccionado: {model_analysis['selected_name']}")
    print("Metricas de test:", model_analysis["test_metrics"])


if __name__ == "__main__":
    main()