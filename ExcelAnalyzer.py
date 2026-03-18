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
        "Tiendas online",
        [
            ("vinted", "Vinted"),
            ("amazon", "Amazon"),
        ],
    ),
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
            (("supermerca", "supermarket", "minimarket", "super", "merca"), "Supermercado generico"),
        ],
    ),
    (
        "Tiendas fisicas",
        [
            ("tiendanimal", "Tiendanimal"),
            (("estanco", "tabac"), "Estanco"),
            ("farmacia", "Farmacia"),
            ("el corte ingles", "El Corte Ingles"),
            ("clinica dental", "Clinica Dental"),
            ("humana", "Humana"),
            ("decathlon", "Decathlon"),
            (("agapea", "libreria", "biblioteca"), "Libreria"),
            ("ikea", "Ikea"),
            ("xocolat", "Xocolat CDs"),
            ("sapporet", "Sapporet Vapers"),
            (("bazar", "xixi xu"), "Bazar"),
            ("tezenis", "Tezenis"),
            ("festival pa", "Festival Pa"),

            # Palabras clave que pertenecen a tiendas físicas
            ("monchis", "Monchis"),
            ("flor", "Floristeria"),
            (("barber", "peluqueria"), "Barberia Peluqueria"),
            (("fruta", "verdura"), "Fruteria Verduleria"),

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
        ],
    ),
    (
        "Restaurantes",
        [
            ("bk", "Burger King"),
            ("vicio", "Vicio"),
            ("kebab", "Kebab"),
            ("mcdonalds", "McDonalds"),
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
            (("asia", "susi", "sushi", "ramen", "wok"), "Restaurante Asiatico"),
            (("taco", "taquero"), "Restaurante Mexicano"),
            (("restaurante",), "Restaurante generico"),

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
            (("spotify", "spoty"), "Spotify"),
            ("renting tec", "Renting Movil"),
            ("liquidacion por bonificacion", "Bonificacion Movil"),
        ],
    ),
    (
        "Varios",
        [
            ("cajero", "Cajero"),
            ("smap ora", "SMAP Ora"),
            ("devolucion", "Devolucion"),
            ("salario", "Salario"),
            ("lavanderia", "Lavanderia"),
            ("vendin", "Maquina expendedora"),
            ("wallapop", "Wallapop"),
            (("cumpleaños", "cumple", "regal"), "Cumpleaños"),
            (("sin concepto",), "Sin Concepto"),
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
            ("bus", "Bus"),
            ("parking", "Parking"),
            ("transport", "Transporte generico"),
        ],
    ),
    (
        "Ocio/Fiesta",
        [
            ("factoria de so", "Factoria de So"),
            ("catsmusics jazz", "Catsmusics Jazz"),
            ("atomic garden", "Atomic Garden"),
            ("jovells", "Jovells (Festa?)"),
            ("cinesa", "Cinesa"),
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
            ("quinta", "Quintades"),
            (("biofesta", "bio"), "Biofesta"),
        ],
    ),
    (
        "Perros",
        [
            (("veterina", "cedivet"), "Veterinario"),
            (("pienso", "comida perro"), "Comida Perro"),
            (("cans", "perro", "mascota"), "Cosas de perros"),
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
        ],
    ),
    (
        "Transferencias",
        [
            (("transferencia", "transf"), "Transferencia"),
            ("trade republic", "Trade Republic"),
            (("tricount", "3count"), "Tricount"),
            (("impuesto", "hacienda"), "Impuestos"),
        ],
    ),
    (
        "Seguros",
        [
            ("prima", "Prima"),
            ("generali", "Generali Seguros"),
        ],
    )
]


def categorizar(desc):
    desc = desc.lower()

    for big_category, patterns in CATEGORY_RULES:
        for pattern, label in patterns:
            if isinstance(pattern, str):
                if pattern in desc:
                    return big_category, label
            elif isinstance(pattern, tuple):
                if any(p in desc for p in pattern):
                    return big_category, label
            else:  # regex
                if pattern.search(desc):
                    return big_category, label

    return "Otros", "Otros"


def generate_plots(df: pd.DataFrame):
    """Generate and save charts based on the processed dataframe."""

    plots_dir = Path(__file__).parent / "plots"
    plots_dir.mkdir(exist_ok=True)

    # Ensure dates are sorted for time-series plots
    df_sorted = df.sort_values("fecha")

    # 1) Spend by category (top 20)
    spend_by_cat = (
        df_sorted.groupby("categoria")["gasto"].sum().sort_values(ascending=False).head(20)
    )
    plt.figure(figsize=(10, 6))
    sns.barplot(x=spend_by_cat.values, y=spend_by_cat.index, palette="viridis")
    plt.title("Gasto total por categoría (top 20)")
    plt.xlabel("Gasto (€)")
    plt.tight_layout()
    plt.savefig(plots_dir / "gasto_por_categoria.png")
    plt.close()

    # 2) Balance timeline
    if "saldo" in df_sorted.columns and df_sorted["saldo"].notna().any():
        plt.figure(figsize=(12, 5))
        sns.lineplot(x="fecha", y="saldo", data=df_sorted, marker="o")
        plt.title("Evolución del saldo")
        plt.xlabel("Fecha")
        plt.ylabel("Saldo")
        plt.tight_layout()
        plt.savefig(plots_dir / "saldo_timeline.png")
        plt.close()

    # 3) Biggest individual spends
    biggest_spends = df_sorted.nlargest(15, "gasto")[["descripcion", "categoria", "gasto"]]
    plt.figure(figsize=(10, 6))
    sns.barplot(
        x="gasto",
        y="descripcion",
        data=biggest_spends,
        palette="magma",
    )
    plt.title("Mayores gastos individuales (top 15)")
    plt.xlabel("Gasto (€)")
    plt.ylabel("Descripción")
    plt.tight_layout()
    plt.savefig(plots_dir / "mayores_gastos.png")
    plt.close()

    # 4) Heatmap: gasto por categoría y mes
    df_sorted["mes"] = df_sorted["fecha"].dt.to_period("M").astype(str)
    pivot = (
        df_sorted[df_sorted["gasto"] > 0]
        .pivot_table(
            index="categoria",
            columns="mes",
            values="gasto",
            aggfunc="sum",
            fill_value=0,
        )
    )
    if not pivot.empty:
        pivot = pivot.loc[pivot.sum(axis=1).sort_values(ascending=False).index]

    if not pivot.empty:
        plt.figure(figsize=(12, max(6, len(pivot) * 0.25)))
        sns.heatmap(pivot, cmap="YlOrRd", linewidths=0.5)
        plt.title("Gasto por categoría y mes")
        plt.xlabel("Mes")
        plt.ylabel("Categoría")
        plt.tight_layout()
        plt.savefig(plots_dir / "gasto_por_categoria_mes.png")
        plt.close()

        # Normalize per category (row) so each category is compared within itself
        pivot_norm = pivot.div(pivot.max(axis=1), axis=0).fillna(0)
        plt.figure(figsize=(12, max(6, len(pivot_norm) * 0.25)))
        sns.heatmap(
            pivot_norm,
            cmap="YlOrRd",
            linewidths=0.5,
            vmin=0,
            vmax=1,
            annot=pivot.apply(lambda col: col.map(lambda v: f"{v:,.0f}")),
            fmt="",
        )
        plt.title("Gasto por categoría y mes (normalizado por categoría)")
        plt.xlabel("Mes")
        plt.ylabel("Categoría")
        plt.tight_layout()
        plt.savefig(plots_dir / "gasto_por_categoria_mes_normalized.png")
        plt.close()

    # 5) Monthly trend per category (compare each category with itself across months)
    monthly = (
        df_sorted[df_sorted["gasto"] > 0]
        .groupby(["categoria", "mes"])["gasto"]
        .sum()
        .reset_index()
    )
    if not monthly.empty:
        # Keep only top categories by total spend to avoid overcrowded plots
        top_categories = (
            monthly.groupby("categoria")["gasto"].sum().nlargest(10).index
        )
        monthly_top = monthly[monthly["categoria"].isin(top_categories)].copy()

        # Exclude Seguros from this plot (requested)
        monthly_top = monthly_top[monthly_top["categoria"] != "Seguros"].copy()

        # Ensure log-scale works (cannot plot 0)
        monthly_top["gasto"] = monthly_top["gasto"].clip(lower=1)
        monthly_top["mes_dt"] = pd.to_datetime(monthly_top["mes"])

        plt.figure(figsize=(12, 6))
        sns.lineplot(
            data=monthly_top,
            x="mes_dt",
            y="gasto",
            hue="categoria",
            marker="o",
        )
        ax = plt.gca()
        ax.set_yscale("symlog", linthresh=250, linscale=1)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda y, _: f"{int(y):,}"))
        ax.yaxis.set_minor_formatter(mticker.NullFormatter())
        ax.yaxis.set_major_locator(mticker.MaxNLocator(nbins=10, integer=True))

        # Annotate the latest value for each category at the right edge
        last_month = monthly_top["mes_dt"].max()
        last_values = (
            monthly_top[monthly_top["mes_dt"] == last_month]
            .set_index("categoria")["gasto"]
        )
        for cat, line in zip(monthly_top["categoria"].unique(), ax.lines):
            if cat in last_values:
                y = last_values[cat]
                # Normalize y for symlog plotting
                ax.text(
                    last_month,
                    y,
                    f" {int(y):,}",
                    va="center",
                    ha="left",
                    fontsize=8,
                )

        plt.title("Gasto mensual por categoría (top 10, sin Seguros)")
        plt.xlabel("Mes")
        plt.ylabel("Gasto (€)")
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig(plots_dir / "gasto_mensual_por_categoria.png")
        plt.close()

    print(f"\nCharts written to: {plots_dir.resolve()}")


def read_excels(__file__):
    excel_dir = Path(__file__).parent / "excels"
    excel_files = list(excel_dir.glob("*.xls"))
    print(f"Found {len(excel_files)} Excel files:")
    for file in excel_files:
        print(f"  - {file}")

    dfs = [pd.read_excel(file).iloc[7:] for file in excel_files]
    return dfs

dfs = read_excels(__file__)

if not dfs:
    print("No Excel files found in the specified directory.")
else:
    df = pd.concat(dfs, ignore_index=True)
    df.drop_duplicates(inplace=True)

    df.drop(df.columns[0], axis=1, inplace=True)

    df.columns = ["fecha", "descripcion", "importe", "saldo"]
    df["fecha"] = pd.to_datetime(df["fecha"], format="%d/%m/%Y")

    df[["categoria", "subcategoria"]] = pd.DataFrame(
        df["descripcion"].apply(categorizar).tolist(), index=df.index
    )
    df = df[df["categoria"] != "Transferencias"].reset_index(drop=True)

    # normalize numeric columns
    df["importe"] = pd.to_numeric(df["importe"], errors="coerce")
    df["saldo"] = pd.to_numeric(df["saldo"], errors="coerce")

    # Build expense / income helpers
    df["gasto"] = -df["importe"].where(df["importe"] < 0, 0)
    df["ingreso"] = df["importe"].where(df["importe"] > 0, 0)

    print("Shape of the merged dataframe:", df.shape)

    generate_plots(df)

