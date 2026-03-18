from pathlib import Path
import pandas as pd
import re

try:
    import matplotlib.pyplot as plt
    import matplotlib.ticker as mticker
    import seaborn as sns
except ImportError:
    plt = None
    sns = None

pd.set_option('display.max_colwidth', None)

CATEGORY_RULES = [
    (
        "Supermercados",
        [
            ("mercadona", "Mercadona"),
            ("coaliment", "Coaliment"),
            ("alcampo ", "Alcampo"),
            ("eroski", "Eroski"),
            ("suma", "Suma"),
            ("syp", "Syp"),
            ("mueller", "Muller (?)"),
            ("bip", "Bip"),
            ("lidl", "Lidl"),
            ("aldi", "Aldi"),
            ("condis", "Condis"),
            (re.compile(r"\bdia\b"), "Dia supermercado"),
            ("la sirena", "La Sirena super"),
            ("granier", "Granier super"),
            (("supermerca", "supermarket", "minimarket", "super", "merca", "hiper"), "Supermercado generico"),
            ("vendin", "Maquina expendedora"),
            (("bazar", "xixi xu"), "Bazar"),
            (("estanco", "tabac"), "Estanco"),
            ("ramis sencelles", "Ramis Sencelles"),
        ],
    ),
    (
        "Tiendas fisicas",
        [
            ("el corte ingles", "El Corte Ingles"),
            ("decathlon", "Decathlon"),
            ("ikea", "Ikea"),
            ("sapporet", "Sapporet Vapers"),
            ("tezenis", "Tezenis"),
            ("festival pa", "Festival Pa"),
            ("fundas", "Fundas"),

            # Palabras clave que pertenecen a tiendas físicas
            ("monchis", "Monchis"),
            ("flor", "Floristeria"),
            (("barber", "peluqueria"), "Barberia Peluqueria"),
            (("fruta", "verdura", "fruteria"), "Fruteria Verduleria"),
        ],
    ),
    (
        "Moda",
        [
            ("pull and bear", "Pull and Bear"),
            ("zara", "Zara"),
            ("springfield", "Springfield"),
            ("bershka", "Bershka"),
            ("vans", "Vans"),
            ("urban outfitters", "Urban Outfitters"),
            ("humana", "Humana"),
            ("vinted", "Vinted"),
            (re.compile(r"\bzara\b"), "Zara"),
        ],
    ),
    (
        "Restaurantes",
        [
            ("bk", "Burger King"),
            ("vicio", "Vicio"),
            ("kebab", "Kebab"),
            (("mcdonalds", "macdonalds"), "McDonalds"),
            ("safra son caste", "Es Safra"),
            ("el perro lechero", "El Perro Lechero"),
            ("la bufala", "La Bufala"),
            ("ca'n cannoli", "Can Cannoli"),
            ("natur", "Natur Poke"),
            ("glovo", "Glovo"),
            ("uber eats", "Uber Eats"),
            ("celler sa font", "Celler Sa Font"),

            # Palabras clave que pertenecen a restaurantes
            (("llonguet", "cena", "cenita", "pizz", "hambur", "comida", "sopar", "pisa"), "Cena"),
            (("asia", "susi", "sushi", "ramen", "wok", "suchi"), "Restaurante Asiatico"),
            (("taco", "taquero"), "Restaurante Mexicano"),
            (("restaurant",), "Restaurante generico"),
        ],
    ),
    (
        "Bar/Cafeterias",
        [
            (("coffee", "cafe", "la fabrique"), "Cafetería"),
            ("starbucks", "Starbucks"),
            ("Montaditos", "100 Montaditos"),
            ("lacerve", "Lacerve"),
            ("es trasto", "Es Trasto"),
            ("open marratxi", "Open Marratxi"),
            ("100 montaditos", "100 Montaditos"),
            ("can joan de saigo", "Can Joan de Saigo"),
            ("romani 41", "Es Romani"),
            ("triangolo", "Triangolo pizza"),
            ("city arms", "Bar City Arms"),
            ("petit bistro", "Bar Petit Bistro"),
            ("zumardi", "Bar Zumardi"),
            ("ligabue ferries", "Bar ferry"),
            ("la resistencia", "Bar La Resistencia"),

            # Palabras clave que pertenecen a bares/cafeterías
            (("cola", "refresco"), "refresco"),
            ("brava", "Picada"),
            (("cerve", "birra", "pomada", "alcohol", "caña", "vino"), "Melopeas"),
            ("cantina", "Cantina"),
            (re.compile(r"\bbar\b"), "Bar generico"),
        ],
    ),
    (
        "Mensualidades",
        [
            ("asociacion espanola contra el cancer", "Asociacion Cancer"),
            ("fundacion diabetes cero", "Fundacion Diabetes Cero"),
            (("spotify", "spoty", "dpoti"), "Spotify"),
            ("renting tec", "Renting Movil"),
            ("liquidacion por bonificacion", "Bonificacion Movil"),
        ],
    ),
    (
        "Varios",
        [
            (("cajero", "efectivo", "efectiu"), "Cajero"),
            ("smap ora", "SMAP Ora"),
            ("devolucion", "Devolucion"),
            ("lavanderia", "Lavanderia"),
            ("wallapop", "Wallapop"),
            (("tricount", "3count"), "Tricount"),
            ("amazon", "Amazon"),
        ],
    ),
    (
        "Gasolineras",
        [
            ("plenergy", "Plenergy"),
            ("autonet", "Autonet Oil"),
            ("v2", "V2"),
            ("e.s.santa", "Shell (?)"),
            ("olleries power", "Olleries Power"),
            (("g binissalem", "estacion serv"), "Gasolinera Generica"),
            ("repsol", "Repsol"),
        ],
    ),
    (
        "Transportes/Viajes",
        [
            (("pesa, san sebastian", "transportes pes"), "Pesa"),
            ("vueling", "Vueling"),
            ("ryanair", "Ryanair"),
            ("air europa", "Air Europa"),
            ("balearia", "Balearia"),
            ("iberia", "Iberia fly"),
            ("servicios selecta e", "Metro"),
            (("compañia del tr",), "Tren"),
            ("renfe", "Renfe"),
            (("taxi", "tasi"), "Taxi"),
            ("hotel", "Hotel"),
            ("airbnb", "Airbnb"),

            # Palabras clave que pertenecen a transportes/viajes
            ("metro", "Metro"),
            ("parking", "Parking"),
            ("transport", "Transporte generico"),
            (re.compile(r"\bbus\b"), "Bus"),
        ],
    ),
    (
        "Fiesta",
        [
            ("factoria de so", "Factoria de So"),
            ("catsmusics jazz", "Catsmusics Jazz"),
            ("atomic garden", "Atomic Garden"),
            ("jovells", "Festes de Santa Maria"),
            ("es pou", "Es Pou"),
            ("dabadaba", "Dabadaba"),
            ("cafe milano", "Cafe Milano"),
            ("razzmatazz", "Razzmatazz"),
            ("l ovella", "Ovella Negra"),
            ("sa lluna", "Sa Lluna"),
            ("bigfoot", "Bigfoot"),
            ("apolo", "Apolo"),
            ("perku", "Perku Pub"),
            ("coyote", "Coyote Pub"),
            ("quintad", "Quintades"),
            ("biofesta", "Biofesta"),
            (re.compile(r"\bbio\b"), "Biofesta"),
            (("festa", "paylogic"), "Fiesta Generica"),
        ],
    ),
    (
        "Ocio",
        [
            ("cinesa", "Cinesa"),
            (("agapea", "libreria", "biblioteca"), "Libreria"),
            ("xocolat", "Xocolat CDs"),
            ("bowlin", "Bolos"),
        ]
    ), 
    (
        "Perros",
        [
            (("veterina", "cedivet"), "Veterinario"),
            (("pienso", "comida perro"), "Comida Perro"),
            (("cans", "perro", "mascota"), "Cosas de perros"),
            ("tiendanimal", "Tiendanimal"),

        ],
    ),
    (
        "Salud",
        [
            (("farmacia", "vicente tomas m"), "Farmacia"),
            ("clinica dental", "Clinica Dental"),
            ("tricologia", "Tricologia"),
        ],
    ),
    (
        "Deportes",
        [
            (("monobloc", "in rock", "climb", "rocodrom", "escalada"), "Escalada"),
            (("snow", "esqui", "masella"), "Snowboard"),
            ("padel", "Padel"),
        ],
    ),
    (
        "Coches",
        [
            (("taller", "mecanico"), "Taller"),
            (("coche", "coch"), "Cosas de Coches"),
            ("autodoc", "Autodoc"),
        ],
    ),
    (
        "Conciertos",
        [
            (("dicefm", "dice.fm"), "Dice"),
            ("eventim", "Eventim"),
            (("ticketmaster",), "Ticketmaster"),
            ("ticketvip", "Ticketvip"),
            ("crocantickets", "Crocantickets"),
            ("fourvenues", "Fourvenues"),
            ("tecteltic", "Tecteltic entradas"),
        ],
    ),
    (
        "Inversiones",
        [
            ("finetix limited", "Finetix"),
            ("inversion", "Inversiones"),
            ("sp 500", "S&P 500"),
            ("cryptos", "Cryptos"),
        ],
    ),
    (
        "Transferencias",
        [
            (("Transferencia De Axis Data", "salario"), "Nomina"),
            ("transf", "Transferencia"),
            ("trade republic", "Trade Republic"),
            (("impuesto", "hacienda"), "Impuestos"),
        ],
    ),
    (
        "Seguros",
        [
            ("prima", "Prima"),
            ("generali", "Generali Seguros"),
        ],
    ),
    (
        "Sin Concepto",
        [
            ("sin concepto", "Sin Concepto"),
        ],
    ),
]


def categorizar(desc: str, importe: float | None = None) -> tuple[str, str]:
    """Assign a transaction description to a category + subcategory.

    Special rule: if the description matches the "cumple"/"regal" keywords:
    - Positive importe => Categoria: Varios, Subcategoria: Cumpleaños
    - Negative importe => Categoria: Varios, Subcategoria: Regalo
    """

    desc = str(desc).lower()

    cumple_keywords = ("cumple", "regal", "reyes")
    if any(k in desc for k in cumple_keywords) and importe is not None:
        if importe > 0:
            return "Cumples/Reyes", "Regalo"
        else:
            return "Varios", "Regalo"

    for big_category, patterns in CATEGORY_RULES:
        for item in patterns:
            # Support both (pattern, label) and (pattern1, pattern2..., label)
            if isinstance(item, (tuple, list)) and len(item) >= 2:
                *pattern_values, label = item
                pattern = pattern_values[0] if len(pattern_values) == 1 else tuple(pattern_values)
            else:
                continue

            if isinstance(pattern, str):
                if pattern in desc:
                    return big_category, label
            elif isinstance(pattern, (tuple, list)):
                if any(p in desc for p in pattern):
                    return big_category, label
            else:  # regex
                if pattern.search(desc):
                    return big_category, label

    return "Otros", "Otros"


def _plot_heatmap(
    pivot: pd.DataFrame,
    title: str,
    output_path: Path,
    cmap: str = "YlOrRd",
    vmin: float | None = 0,
    vmax: float | None = 1,
):
    """Plot a normalized heatmap from a pivot table."""

    if pivot.empty:
        return

    # Normalize each row (category) by its maximum absolute value so that
    # categories with only negative net spending still show a gradient.
    row_max = pivot.abs().max(axis=1).replace(0, 1)
    pivot_norm = pivot.div(row_max, axis=0).fillna(0)

    plt.figure(figsize=(12, max(6, len(pivot_norm) * 0.25)))
    sns.heatmap(
        pivot_norm,
        cmap=cmap,
        linewidths=0.5,
        vmin=vmin,
        vmax=vmax,
        annot=pivot.apply(lambda col: col.map(lambda v: f"{v:,.0f}")),
        fmt="",
    )
    plt.title(title)
    plt.xlabel("Mes")
    plt.ylabel("Categoría")
    plt.tight_layout()
    plt.savefig(output_path)
    plt.close()


def generate_plots(df: pd.DataFrame):
    """Generate and save charts based on the processed dataframe."""

    plots_dir = Path(__file__).parent / "plots"
    plots_dir.mkdir(exist_ok=True)

    # Ensure dates are sorted for time-series plots
    df_sorted = df.sort_values("fecha")
    df_sorted["mes"] = df_sorted["fecha"].dt.to_period("M").astype(str)

    # Base spending pivot (gasto only)
    gasto_pivot = (
        df_sorted[df_sorted["gasto"] > 0]
        .pivot_table(
            index="categoria",
            columns="mes",
            values="gasto",
            aggfunc="sum",
            fill_value=0,
        )
    )
    gasto_pivot = gasto_pivot.loc[gasto_pivot.sum(axis=1).sort_values(ascending=False).index]

    _plot_heatmap(
        gasto_pivot,
        "Gasto por categoría y mes (normalizado por categoría)",
        plots_dir / "spending_heatmap.png",
    )

    # Net spending (gasto - ingreso)
    net_pivot = (
        df_sorted
        .pivot_table(
            index="categoria",
            columns="mes",
            values=["gasto", "ingreso"],
            aggfunc="sum",
            fill_value=0,
        )
    )
    net_pivot = net_pivot["gasto"] - net_pivot["ingreso"]
    net_pivot = net_pivot.fillna(0)
    net_pivot = net_pivot.loc[net_pivot.sum(axis=1).sort_values(ascending=False).index]

    _plot_heatmap(
        net_pivot,
        "Gasto neto por categoría y mes (gasto - ingreso)",
        plots_dir / "spending_heatmap_net.png",
        cmap="RdYlGn_r",
        vmin=-1,
        vmax=1,
    )

    # 2) Biggest individual spends
    biggest_spends = df_sorted.nlargest(15, "gasto")[["descripcion", "categoria", "gasto"]]
    plt.figure(figsize=(10, 6))
    sns.barplot(
        x="gasto",
        y="descripcion",
        data=biggest_spends,
        hue="descripcion",
        palette="magma",
        dodge=False,
    )
    plt.legend([], [], frameon=False)
    plt.title("Mayores gastos individuales (top 15)")
    plt.xlabel("Gasto (€)")
    plt.ylabel("Descripción")
    plt.tight_layout()
    plt.savefig(plots_dir / "biggest_individual_spends.png")
    plt.close()

    print(f"\nCharts written to: {plots_dir.resolve()}")


EXCEL_GLOB_PATTERN = "*.xls"
OUTPUT_FILENAME = "processed_transactions.xlsx"


def _sanitize_sheet_name(name: str) -> str:
    """Normalize a string to a valid Excel sheet name."""

    invalid = r"[]:*?/\\"
    clean = "".join(c for c in str(name) if c not in invalid)
    return clean[:31]


def _format_sheet(
    worksheet,
    data_frame: pd.DataFrame,
    workbook,
    date_fmt,
    date_bold_center_fmt,
):
    """Apply consistent formatting to a worksheet containing `data_frame`."""

    # Header formatting
    header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9D9D9", "align": "center"})
    for col_num, value in enumerate(data_frame.columns):
        worksheet.write(0, col_num, value, header_fmt)

    # Base cell formats
    center_fmt = workbook.add_format({"align": "center"})
    left_fmt = workbook.add_format({"align": "left"})
    bold_center_fmt = workbook.add_format({"bold": True, "align": "center"})

    for col_num, col in enumerate(data_frame.columns):
        max_len = max(data_frame[col].astype(str).map(len).max(), len(col)) + 2

        if col == "fecha":
            fmt = date_bold_center_fmt
        elif col == "categoria":
            fmt = bold_center_fmt
        elif col == "descripcion":
            fmt = left_fmt
        else:
            fmt = center_fmt

        worksheet.set_column(col_num, col_num, max_len, fmt)

    # Conditional formatting for importe: green for positive, red for negative
    if "importe" in data_frame.columns:
        importe_col = data_frame.columns.get_loc("importe")
        fmt_pos = workbook.add_format({"font_color": "green"})
        fmt_neg = workbook.add_format({"font_color": "red"})
        worksheet.conditional_format(
            1,
            importe_col,
            len(data_frame),
            importe_col,
            {"type": "cell", "criteria": ">", "value": 0, "format": fmt_pos},
        )
        worksheet.conditional_format(
            1,
            importe_col,
            len(data_frame),
            importe_col,
            {"type": "cell", "criteria": "<", "value": 0, "format": fmt_neg},
        )

    # Force fecha to be displayed in dd/mm/yyyy, bold, and centered by rewriting the datetime cells
    if "fecha" in data_frame.columns:
        fecha_col = data_frame.columns.get_loc("fecha")
        for row_idx, value in enumerate(data_frame["fecha"], start=1):
            if pd.notna(value):
                worksheet.write_datetime(
                    row_idx, fecha_col, value.to_pydatetime(), date_bold_center_fmt
                )

    # Zebra striping for easier reading
    zebra_fmt = workbook.add_format({"bg_color": "#F2F2F2"})
    worksheet.conditional_format(
        1,
        0,
        len(data_frame),
        len(data_frame.columns) - 1,
        {
            "type": "formula",
            "criteria": "=MOD(ROW(),2)=0",
            "format": zebra_fmt,
        },
    )


def read_excels(excel_dir: Path | str, pattern: str = EXCEL_GLOB_PATTERN) -> list[pd.DataFrame]:
    """Load all Excel files matching a glob pattern and return a list of DataFrames."""

    excel_dir = Path(excel_dir)
    excel_files = list(excel_dir.glob(pattern))
    print(f"Found {len(excel_files)} Excel files:")
    for file in excel_files:
        print(f"  - {file}")

    return [pd.read_excel(file).iloc[7:] for file in excel_files]


def prepare_transactions(df: pd.DataFrame) -> pd.DataFrame:
    """Clean and enrich the transactions DataFrame."""

    df = df.copy()

    # Remove redundant first column introduced by the source Excel files
    df.drop(df.columns[0], axis=1, inplace=True)

    df.columns = ["fecha", "descripcion", "importe", "saldo"]
    df["fecha"] = pd.to_datetime(df["fecha"], format="%d/%m/%Y", errors="coerce")

    df[["categoria", "subcategoria"]] = pd.DataFrame(
        df.apply(lambda row: categorizar(row["descripcion"], row["importe"]), axis=1).tolist(),
        index=df.index,
    )
    df = df[df["categoria"] != "Transferencias"].reset_index(drop=True)

    # normalize numeric columns
    df["importe"] = pd.to_numeric(df["importe"], errors="coerce")
    df["saldo"] = pd.to_numeric(df["saldo"], errors="coerce")

    # Build expense / income helpers
    df["gasto"] = -df["importe"].where(df["importe"] < 0, 0)
    df["ingreso"] = df["importe"].where(df["importe"] > 0, 0)

    return df


def save_transactions_xlsx(df: pd.DataFrame, out_xlsx: Path):
    """Save transactions to an Excel file with nice formatting."""

    df_export = df.sort_values("fecha", ascending=False).reset_index(drop=True)
    # df_export = df_sorted.drop(columns=["gasto", "ingreso"], errors="ignore")

    try:
        with pd.ExcelWriter(
            out_xlsx,
            engine="xlsxwriter",
            date_format="dd/mm/yyyy",
            datetime_format="dd/mm/yyyy",
        ) as writer:
            workbook = writer.book

            # Format helpers
            date_fmt = workbook.add_format({"num_format": "dd/mm/yyyy", "align": "center"})
            date_bold_center_fmt = workbook.add_format(
                {"num_format": "dd/mm/yyyy", "align": "center", "bold": True}
            )

            # Full sheet with all transactions
            full_sheet_name = "Todos" if "Todos" not in writer.sheets else "Todos_1"
            df_export.to_excel(writer, index=False, sheet_name=full_sheet_name)
            _format_sheet(
                writer.sheets[full_sheet_name],
                df_export,
                workbook,
                date_fmt,
                date_bold_center_fmt,
            )

            # Per-category sheets
            seen = {full_sheet_name}
            for category, group in df_export.groupby("categoria"):
                if not str(category).strip():
                    continue

                sheet_name = _sanitize_sheet_name(category) or "SinCat"
                base = sheet_name
                suffix = 1
                while sheet_name in seen:
                    sheet_name = f"{base}_{suffix}"
                    suffix += 1
                seen.add(sheet_name)

                group.to_excel(writer, index=False, sheet_name=sheet_name)
                _format_sheet(
                    writer.sheets[sheet_name],
                    group,
                    workbook,
                    date_fmt,
                    date_bold_center_fmt,
                )

        print(f"Formatted Excel written to: {out_xlsx.resolve()}")
    except Exception as e:
        print(f"Could not write Excel file: {e}")


def main():
    excel_dir = Path(__file__).parent / "excels"
    dfs = read_excels(excel_dir)

    if not dfs:
        print("No Excel files found in the specified directory.")
        return

    df = pd.concat(dfs, ignore_index=True)
    df.drop_duplicates(inplace=True)

    df = prepare_transactions(df)

    print("Shape of the merged dataframe:", df.shape)

    out_xlsx = Path(__file__).parent / OUTPUT_FILENAME
    save_transactions_xlsx(df, out_xlsx)

    generate_plots(df)


if __name__ == "__main__":
    main()

