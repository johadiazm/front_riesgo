import pandas as pd
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import seaborn as sns
from fpdf import FPDF
import os
import plotly.express as px
import matplotlib
matplotlib.use('Agg')

# grafico_estadocivil
def grafico_estadocivil(df,img_path):
    df_grouped = df["4. ESTADO CIVIL"].value_counts().reset_index()
    df_grouped.columns = ["Estado Civil", "Cantidad"]

    # Crear gráfico de pastel en 3D
    fig = go.Figure(data=[go.Pie(labels=df_grouped["Estado Civil"], values=df_grouped["Cantidad"], hole=0.3)])

    fig.update_traces(pull=[0.1] * len(df_grouped), textinfo="percent", marker=dict(line=dict(color="black", width=1)))
    fig.update_layout(title="Estado Civil ",scene_camera=dict(eye=dict(x=5, y=2, z=0.3))
                      ,width=300, height=300)

    # Guardar gráfico como imagen
    fig.write_image(img_path)


# grafico_tipo_sexo
def grafico_tipo_sexo(df, img_path):
    df_grouped = df["2. SEXO"].value_counts()
    # Definir colores y propiedades de borde
    colors = plt.cm.Paired.colors[:len(df_grouped)]  # Colores personalizados
    wedge_properties = {"edgecolor": "black", "linewidth": 0.5}  # Bordes 
    # Crear el gráfico de pastel con bordes
    plt.figure(figsize=(2, 2))
    wedges, texts, autotexts = plt.pie(
        df_grouped, 
        labels=df_grouped.index, 
        autopct='%1.1f%%', 
        startangle=90, 
        colors=colors, 
        wedgeprops=wedge_properties
    )
    # Agregar la leyenda
    plt.legend(wedges, df_grouped.index, title="SEXO", loc="lower left", bbox_to_anchor=(1, 0, 0.5, 1))
    plt.title("TIPO DE SEXO", fontsize=8)
    plt.savefig(img_path)
    plt.close()

# grafico_tipo_vivienda
def grafico_tipo_vivienda(df, img_path):
    df_grouped = df["9. TIPO DE VIVIENDA"].value_counts()
    colors = plt.cm.Paired.colors[:len(df_grouped)]  # Colores personalizados

    plt.figure(figsize=(2, 2))
    wedges, texts, autotexts = plt.pie(
        df_grouped,
        labels=df_grouped.index,
        autopct='%1.1f%%',
        startangle=90,
        colors=colors
    )
    plt.legend(wedges, df_grouped.index, title="TIPO DE VIVIENDA", loc="lower left", bbox_to_anchor=(1, 0, 0.5, 1))
    plt.title("TIPO DE VIVIENDA")
    plt.savefig(img_path)
    plt.close()

# grafico_personas_a_cargo
def grafico_personas_a_cargo(df, img_path):
    df_grouped = df["10. NUMERO DE PERSONAS QUE DEPENDEN ECONOMICAMENTE DE USTED"].value_counts()
    colors = plt.cm.Paired.colors[:len(df_grouped)]
    plt.figure(figsize=(2, 2))
    wedges, texts, autotexts = plt.pie(
        df_grouped,
        labels=df_grouped.index,
        autopct='%1.1f%%',
        startangle=90,
        colors=colors
    )
    plt.legend(
        wedges,
        df_grouped.index,
        title="Número de Personas a Cargo",
        loc="center left",
        bbox_to_anchor=(1, 0, 0.5, 1)
    )
    plt.title("Distribución de Personas Económicamente a Cargo")
    plt.savefig(img_path)
    plt.close()

# grafico_tipo_contrato
def grafico_tipo_contrato(df, img_path):
    df_grouped = df["17. TIPO DE CONTRATO ACTUAL"].value_counts()
    colors = plt.cm.Paired.colors[:len(df_grouped)]  # Colores personalizados

    plt.figure(figsize=(2, 2))
    wedges, texts, autotexts = plt.pie(
        df_grouped,
        labels=df_grouped.index,
        autopct='%1.1f%%',
        startangle=90,
        colors=colors
    )
    plt.legend(
        wedges,
        df_grouped.index,
        title="TIPO DE CONTRATO",
        loc="lower left",
        bbox_to_anchor=(1, 0, 0.5, 1)
    )
    plt.title("TIPO DE CONTRATO")
    plt.savefig(img_path)
    plt.close()

# funcion para grafico de nivel generacional falta


# grafico_escolaridad
def grafico_escolaridad(df, img_path):
    df_grouped = df["5. NIVEL DE ESTUDIOS"].value_counts()
    total = df_grouped.sum()
    plt.figure(figsize=(4, 3))
    ax = sns.barplot(
        x=df_grouped.index,
        y=df_grouped.values,
        palette="viridis"
    )
    plt.title("Distribución de Nivel de Estudios")
    plt.xlabel("Nivel de Estudios")
    plt.ylabel("Número de Personas")
    plt.xticks(rotation=0, fontsize=8)
    for i, v in enumerate(df_grouped.values):
        porcentaje = v / total * 100
        ax.text(i, v * 0.5, f"{porcentaje:.1f}%", ha='center', va='bottom', fontsize=9, color='black')

    plt.tight_layout()
    plt.savefig(img_path)
    plt.close()

# grafico_estrato
def grafico_estrato(df, img_path):
    df_grouped = df["8. ESTRATO SOCIAL"].value_counts()
    colors = plt.cm.Paired.colors[:len(df_grouped)]
    plt.figure(figsize=(2, 2))
    wedges, texts, autotexts = plt.pie(
        df_grouped,
        labels=df_grouped.index,
        autopct='%1.1f%%',
        startangle=90,
        colors=colors
    )
    plt.legend(
        wedges,
        df_grouped.index,
        title="Estrato",
        loc="center left",
        bbox_to_anchor=(1, 0, 0.5, 1)
    )
    plt.title("TIPO ESTRATO")
    plt.savefig(img_path)
    plt.close()




# grafico_riesgos_transformados
def grafico_liderazgo_y_relaciones_sociales_tipo_A(df, img_path):
    # --- Crear columnas transformadas si no existen ---
    if 'Características del liderazgo transformado' not in df.columns:
        columnas_a_sumar = [
            "Mi jefe me da instrucciones claras",
            "Mi jefe ayuda a organizar mejor el trabajo",
            "Mi jefe tiene en cuenta mis puntos de vista y opiniones",
            "Mi jefe me anima para hacer mejor mi trabajo",
            "Mi jefe distribuye las tareas de forma que me facilita el trabajo",
            "Mi jefe me comunica a tiempo la información relacionada con el trabajo",
            "La orientación que me da mi jefe me ayuda a hacer mejor el trabajo",
            "Mi jefe me ayuda a progresar en el trabajo",
            "Mi jefe me ayuda a sentirme bien en el trabajo",
            "Mi jefe ayuda a solucionar los problemas que se presentan en el trabajo",
            "Siento que puedo confiar en mi jefe",
            "Mi jefe me escucha cuando tengo problemas de trabajo",
            "Mi jefe me brinda su apoyo cuando lo necesito"
        ]
        df['Características del liderazgo transformado'] = (
            df[columnas_a_sumar].apply(
                lambda row: row.sum() if not row.isnull().any() else None, axis=1
            ) / 52 * 100
        )

    if 'Relaciones sociales en el trabajo transformada' not in df.columns:
        columnas_a_sumar = [
            'Me agrada el ambiente de mi grupo de trabajo',
            'En mi grupo de trabajo me tratan de forma respetuosa',
            'Siento que puedo confiar en mis compañeros de trabajo',
            'Me siento a gusto con mis compañeros de trabajo',
            'En mi grupo de trabajo algunas personas me maltratan',
            'Entre compañeros solucionamos los problemas de forma respetuosa',
            'Hay integración en mi grupo de trabajo',
            'Mi grupo de trabajo es muy unido',
            'Las personas en mi trabajo me hacen sentir parte del grupo',
            'Es fácil poner de acuerdo al grupo para hacer el trabajo',
            'Mis compañeros de trabajo me ayudan cuando tengo dificultades',
            'En mi trabajo las personas nos apoyamos unos a otros',
            'Algunos compañeros de trabajo me escuchan cuando tengo problemas',
            'Cuando tenemos que realizar trabajo de grupo los compañeros colaboran'
        ]
        df['Relaciones sociales en el trabajo transformada'] = (
            df[columnas_a_sumar].apply(
                lambda row: row.sum() if not row.isnull().any() else None, axis=1
            ) / 56 * 100
        )

    if 'Retroalimentación del desempeño transformada' not in df.columns:
        columnas_a_sumar = [
            'Me informan sobre lo que hago bien en mi trabajo',
            'Me informan sobre lo que debo mejorar en mi trabajo',
            'La información que recibo sobre mi rendimiento en el trabajo es clara',
            'La forma como evalúan mi trabajo en la empresa me ayuda a mejorar',
            'Me informan a tiempo sobre lo que debo mejorar en el trabajo'
        ]
        df['Retroalimentación del desempeño transformada'] = (
            df[columnas_a_sumar].apply(
                lambda row: row.sum() if not row.isnull().any() else None, axis=1
            ) / 20 * 100
        )

    if 'Relación con los colaboradores transformada' not in df.columns:
        columnas_a_sumar = [
            'Tengo colaboradores que comunican tarde los asuntos de trabajo',
            'Tengo colaboradores que tienen comportamientos irrespetuosos',
            'Tengo colaboradores que dificultan la organización del trabajo',
            'Tengo colaboradores que guardan silencio cuando les piden opiniones',
            'Tengo colaboradores que dificultan el logro de los resultados del trabajo',
            'Tengo colaboradores que expresan de forma irrespetuosa sus desacuerdos',
            'Tengo colaboradores que cooperan poco cuando se necesita',
            'Tengo colaboradores que me preocupan por su desempeño',
            'Tengo colaboradores que ignoran las sugerencias para mejorar su trabajo'
        ]
        df['Relación con los colaboradores transformada'] = (
            df[columnas_a_sumar].apply(
                lambda row: row.sum() if not row.isnull().any() else None, axis=1
            ) / 36 * 100
        )

    # --- Lógica de clasificación y gráfico ---
    def clasificar_riesgo(valor, limites):
        if pd.isnull(valor):
            return None
        if limites["muy_alto"][0] <= valor <= limites["muy_alto"][1]:
            return "Riesgo muy alto"
        elif limites["alto"][0] <= valor <= limites["alto"][1]:
            return "Riesgo alto"
        elif limites["medio"][0] <= valor <= limites["medio"][1]:
            return "Riesgo medio"
        elif limites["bajo"][0] <= valor <= limites["bajo"][1]:
            return "Riesgo bajo"
        elif limites["sin_riesgo"][0] <= valor <= limites["sin_riesgo"][1]:
            return "Sin riesgo"
        else:
            return "No clasificado"

    limites_riesgo = {
        "Características del liderazgo transformado": {
            "muy_alto": (46.3, 100),
            "alto": (30.9, 46.2),
            "medio": (15.5, 30.8),
            "bajo": (3.9, 15.4),
            "sin_riesgo": (0.0, 3.8),
        },
        "Relaciones sociales en el trabajo transformada": {
            "muy_alto": (37.6, 100),
            "alto": (25.1, 37.5),
            "medio": (16.2, 25.0),
            "bajo": (5.5, 16.1),
            "sin_riesgo": (0.0, 5.4),
        },
        "Retroalimentación del desempeño transformada": {
            "muy_alto": (55.1, 100),
            "alto": (40.1, 55.0),
            "medio": (25.1, 40.0),
            "bajo": (10.1, 25.0),
            "sin_riesgo": (0.0, 10.0),
        },
        "Relación con los colaboradores transformada": {
            "muy_alto": (45.3, 100),
            "alto": (33.4, 47.2),
            "medio": (25.1, 33.3),
            "bajo": (14.0, 25.0),
            "sin_riesgo": (0.0, 13.9),
        }
    }

    niveles = ["Riesgo muy alto", "Riesgo alto", "Riesgo medio", "Riesgo bajo", "Sin riesgo"]
    conteo = []
    for dim, limites in limites_riesgo.items():
        clasificaciones = df[dim].dropna().apply(lambda x: clasificar_riesgo(x, limites))
        total = len(clasificaciones)
        for nivel in niveles:
            cantidad = (clasificaciones == nivel).sum()
            porcentaje = (cantidad / total * 100) if total > 0 else 0
            conteo.append({"Dimensión": dim, "Riesgo": nivel, "Porcentaje": porcentaje})

    df_plot = pd.DataFrame(conteo)

    plt.figure(figsize=(6, 4))
    colores = ["#d32f2f", "#fbc02d", "#9b9b9b", "#1976d2", "#388e3c"]
    ax = sns.barplot(
        x="Dimensión", y="Porcentaje", hue="Riesgo",
        data=df_plot, palette=colores, order=list(limites_riesgo.keys()), hue_order=niveles
    )
    for p in ax.patches:
        if p.get_height() > 0:
            ax.annotate(f'{p.get_height():.1f}%',
                        (p.get_x() + p.get_width() / 2., p.get_height()),
                        ha='center', va='bottom', fontsize=9, color='black')
    plt.title("Distribución de Niveles de Riesgo por Dimensión")
    plt.xlabel("Dimensión")
    plt.ylabel("Porcentaje (%)")
    plt.xticks(rotation=15, fontsize=9)
    plt.legend(title="Nivel de Riesgo", fontsize=6)
    plt.tight_layout()
    plt.savefig(img_path)
    plt.close()

def grafico_liderazgo_y_relaciones_sociales_tipo_B(df, img_path):
    # --- Crear columnas transformadas si no existen ---
    if 'Características del liderazgo transformado' not in df.columns:
        columnas_a_sumar = [
            "Mi jefe me da instrucciones claras",
            "Mi jefe ayuda a organizar mejor el trabajo",
            "Mi jefe tiene en cuenta mis puntos de vista y opiniones",
            "Mi jefe me anima para hacer mejor mi trabajo",
            "Mi jefe distribuye las tareas de forma que me facilita el trabajo",
            "Mi jefe me comunica a tiempo la información relacionada con el trabajo",
            "La orientación que me da mi jefe me ayuda a hacer mejor el trabajo",
            "Mi jefe me ayuda a progresar en el trabajo",
            "Mi jefe me ayuda a sentirme bien en el trabajo",
            "Mi jefe ayuda a solucionar los problemas que se presentan en el trabajo",
            "Siento que puedo confiar en mi jefe",
            "Mi jefe me escucha cuando tengo problemas de trabajo",
            "Mi jefe me brinda su apoyo cuando lo necesito"
        ]
        df['Características del liderazgo transformado'] = (
            df[columnas_a_sumar].apply(
                lambda row: row.sum() if not row.isnull().any() else None, axis=1
            ) / 52 * 100
        )

    if 'Relaciones sociales en el trabajo transformada' not in df.columns:
        columnas_a_sumar = [
            'Me agrada el ambiente de mi grupo de trabajo',
            'En mi grupo de trabajo me tratan de forma respetuosa',
            'Siento que puedo confiar en mis compañeros de trabajo',
            'Me siento a gusto con mis compañeros de trabajo',
            'En mi grupo de trabajo algunas personas me maltratan',
            'Entre compañeros solucionamos los problemas de forma respetuosa',
            'Hay integración en mi grupo de trabajo',
            'Mi grupo de trabajo es muy unido',
            'Las personas en mi trabajo me hacen sentir parte del grupo',
            'Es fácil poner de acuerdo al grupo para hacer el trabajo',
            'Mis compañeros de trabajo me ayudan cuando tengo dificultades',
            'En mi trabajo las personas nos apoyamos unos a otros',
            'Algunos compañeros de trabajo me escuchan cuando tengo problemas',
            'Cuando tenemos que realizar trabajo de grupo los compañeros colaboran'
        ]
        df['Relaciones sociales en el trabajo transformada'] = (
            df[columnas_a_sumar].apply(
                lambda row: row.sum() if not row.isnull().any() else None, axis=1
            ) / 56 * 100
        )

    if 'Retroalimentación del desempeño transformada' not in df.columns:
        columnas_a_sumar = [
            'Me informan sobre lo que hago bien en mi trabajo',
            'Me informan sobre lo que debo mejorar en mi trabajo',
            'La información que recibo sobre mi rendimiento en el trabajo es clara',
            'La forma como evalúan mi trabajo en la empresa me ayuda a mejorar',
            'Me informan a tiempo sobre lo que debo mejorar en el trabajo'
        ]
        df['Retroalimentación del desempeño transformada'] = (
            df[columnas_a_sumar].apply(
                lambda row: row.sum() if not row.isnull().any() else None, axis=1
            ) / 20 * 100
        )
        
        df['Relación con los colaboradores transformada'] = (
            df[columnas_a_sumar].apply(
                lambda row: row.sum() if not row.isnull().any() else None, axis=1
            ) / 36 * 100
        )

    # --- Lógica de clasificación y gráfico ---
    def clasificar_riesgo(valor, limites):
        if pd.isnull(valor):
            return None
        if limites["muy_alto"][0] <= valor <= limites["muy_alto"][1]:
            return "Riesgo muy alto"
        elif limites["alto"][0] <= valor <= limites["alto"][1]:
            return "Riesgo alto"
        elif limites["medio"][0] <= valor <= limites["medio"][1]:
            return "Riesgo medio"
        elif limites["bajo"][0] <= valor <= limites["bajo"][1]:
            return "Riesgo bajo"
        elif limites["sin_riesgo"][0] <= valor <= limites["sin_riesgo"][1]:
            return "Sin riesgo"
        else:
            return "No clasificado"

    limites_riesgo = {
        "Características del liderazgo transformado": {
            "muy_alto": (38.6, 100),
             "alto": (25.1, 38.5),
            "medio": (13.6, 25.0),
            "bajo": (3.9, 13.5),
            "sin_riesgo": (0.0, 3.8),
        },
        "Relaciones sociales en el trabajo transformada": {
           "muy_alto": (37.6, 100),
           "alto": (27.2, 37.5),
           "medio": (14.7, 27.1),
           "bajo": (6.4, 14.6),
           "sin_riesgo": (0.0, 6.3),
        },
        "Retroalimentación del desempeño transformada": {
            "muy_alto": (50.1, 100),
            "alto": (30.1, 50.0),
            "medio": (20.1, 30.0),
            "bajo": (5.1, 20.0),
            "sin_riesgo": (0.0, 5.0),
        },
       
    }

    niveles = ["Riesgo muy alto", "Riesgo alto", "Riesgo medio", "Riesgo bajo", "Sin riesgo"]
    conteo = []
    for dim, limites in limites_riesgo.items():
        clasificaciones = df[dim].dropna().apply(lambda x: clasificar_riesgo(x, limites))
        total = len(clasificaciones)
        for nivel in niveles:
            cantidad = (clasificaciones == nivel).sum()
            porcentaje = (cantidad / total * 100) if total > 0 else 0
            conteo.append({"Dimensión": dim, "Riesgo": nivel, "Porcentaje": porcentaje})

    df_plot = pd.DataFrame(conteo)

    plt.figure(figsize=(6,4))
    colores = ["#d32f2f", "#fbc02d", "#9b9b9b", "#1976d2", "#388e3c"]
    ax = sns.barplot(
        x="Dimensión", y="Porcentaje", hue="Riesgo",
        data=df_plot, palette=colores, order=list(limites_riesgo.keys()), hue_order=niveles
    )
    for p in ax.patches:
        if p.get_height() > 0:
            ax.annotate(f'{p.get_height():.1f}%',
                        (p.get_x() + p.get_width() / 2., p.get_height()),
                        ha='center', va='bottom', fontsize=9, color='black')
    plt.title("Distribución de Niveles de Riesgo por Dimensión")
    plt.xlabel("Dimensión")
    plt.ylabel("Porcentaje (%)")
    plt.xticks(rotation=15, fontsize=9)
    plt.legend(title="Nivel de Riesgo", fontsize=6)
    plt.tight_layout()
    plt.savefig(img_path)
    plt.close()


# funcion principal para generar graficos y pdf
def generar_graficos_y_pdf(file_path, filename, reports_dir):
    df = pd.read_excel(file_path)
    graficos = [
        ("Estado civil", grafico_estadocivil),
        ("Sexo", grafico_tipo_sexo),
        ("Tipo de Vivienda", grafico_tipo_vivienda),
        ("Personas a Cargo", grafico_personas_a_cargo),
        ("Tipo de Contrato", grafico_tipo_contrato),
        ("Escolaridad", grafico_escolaridad),
        ("Estrato", grafico_estrato),
        ("Liderazgo y relaciones sociales tipo A", grafico_liderazgo_y_relaciones_sociales_tipo_A),
        ("Liderazgo y relaciones sociales tipo B", grafico_liderazgo_y_relaciones_sociales_tipo_B)
     
    ]

    img_paths = []
    for i, (titulo, funcion) in enumerate(graficos):
        img_path = os.path.join(reports_dir, f"{filename}_grafico_{i}.png")
        funcion(df, img_path)
        img_paths.append((titulo, img_path))

    pdf_path = os.path.join(reports_dir, f"{filename}_informe.pdf")
    pdf = FPDF()
    pdf.alias_nb_pages()
    pdf.set_font("Arial", size=12)
    margen=10
    ancho_pagina = pdf.w - 2 * margen
    #img_width = min(90, pdf.w - 20)  # ancho máximo con margen
    #img_height = 80                  # altura fija para todas las imágenes
    #x_centrada = (pdf.w - img_width) / 2
    #y = 30

    for idx, (titulo, img) in enumerate(img_paths):
        # Si es la primera imagen, crea la portada con logo y título
        if idx == 0:
            pdf.add_page()
            pdf.image("../public/logo.jpg", x=10, y=10, w=40)
            # Título
            pdf.set_xy(0, 20)
            pdf.set_font("Arial", "B", 18)
            pdf.cell(0, 40, "Informe de Riesgo Psicosocial", ln=True, align='C')
            y = 60  # Espacio después del título
        # Si la imagen no cabe, agrega nueva página (sin logo ni título)
        elif y + 90 > pdf.h - 20:
            pdf.add_page()
            y = 30
        # Título del gráfico
        #pdf.set_xy(x_centrada, y - 10)
        #pdf.set_font("Arial", "B", 12)
        #pdf.cell(img_width, 10, titulo, align='C')
        # Imagen
        pdf.set_xy(10, y - 10)
        pdf.cell(0, 10, titulo, align='C')
        pdf.image(img, x=10, y=y)
        y +=90

    pdf.output(pdf_path)
    return f"{filename}_informe.pdf"