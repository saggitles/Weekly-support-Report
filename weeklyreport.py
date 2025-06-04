import os
import pandas as pd
from datetime import datetime
from docx import Document
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
import matplotlib.pyplot as plt
import io

# ----------------------------------------
# Funciones auxiliares para inserciones en Word
# ----------------------------------------
def insert_paragraph_after(paragraph, text=None):
    """
    Inserta un nuevo párrafo justo después de `paragraph` y devuelve el objeto Paragraph creado.
    """
    new_p = OxmlElement('w:p')
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        new_para.add_run(text)
    return new_para

def insert_table_after(doc, paragraph, rows, cols):
    """
    Crea una tabla de tamaño rows x cols al final del documento y luego la mueve
    para que quede justo después de `paragraph`. Devuelve la tabla insertada.
    """
    tbl = doc.add_table(rows=rows, cols=cols)
    body = doc._body._body
    body.remove(tbl._tbl)
    idx = list(body).index(paragraph._p)
    body.insert(idx + 1, tbl._tbl)
    return tbl

def style_table_like_image(table):
    """
    Aplica el estilo de la imagen de muestra a la tabla:
    - Cabecera azul claro
    - Texto en negrita y centrado en la cabecera
    - Bordes en todas las celdas
    """
    # Import required modules
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    
    # Apply light blue header style (first row)
    header_row = table.rows[0]
    for cell in header_row.cells:
        # Set light blue background (#B8CCE4) for header
        shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="B8CCE4"/>')
        cell._element.get_or_add_tcPr().append(shading_elm)
        # Make header text bold and centered
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.bold = True
                
    # Apply borders to all cells in the table
    border_attrs = {"sz": 4, "val": "single", "color": "#000000"}
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(
                cell,
                top=border_attrs,
                bottom=border_attrs,
                left=border_attrs,
                right=border_attrs,
            )

# ----------------------------------------
# 1. --- Cargar datos de Excel (soporte) ---
# ----------------------------------------
excel_path = r'D:\l10 report\Data\tickets (10).xlsx'
if not os.path.isfile(excel_path):
    raise FileNotFoundError(f"No se encontró el archivo de soporte: {excel_path}")

df = pd.read_excel(excel_path)

# 1.1 Normalizar 'Status'
df['Status_norm'] = df['Status'].astype(str).str.strip().str.title()

# 1.2 Normalizar 'Category'
df['Category_norm'] = (
    df['Category']
    .astype(str)
    .str.replace(r'\s*\(.*\)', '', regex=True)
    .str.strip()
    .str.title()
)

# 1.3 Totales y distribuciones (soporte)
total_tickets = len(df)

# 1.3.1 Distribución por Status
status_counts = (
    df['Status_norm']
    .value_counts()
    .reset_index()
    .rename(columns={'index': 'Status', 'Status_norm': 'Count'})
)
status_counts.columns = ['Status', 'Count']
# Filter out "To Do" status
status_counts = status_counts[status_counts['Status'] != 'To Do']
status_counts['Percentage'] = (status_counts['Count'].sum() / total_tickets * 100).round(2)
# Recalculate percentages based on filtered total
status_counts['Percentage'] = (status_counts['Count'] / status_counts['Count'].sum() * 100).round(2)

# 1.3.2 Tickets con más de 10 días abiertos
df['createdAt'] = pd.to_datetime(df['createdAt'])
df['createdAt'] = df['createdAt'].dt.tz_localize(None)
now = datetime.now()
df['DaysOpen'] = (now - df['createdAt']).dt.days
# Filter for tickets with more than 10 days open AND status of "In Progress"
df_old = df[(df['DaysOpen'] > 10) & (df['Status_norm'] == 'In Progress')].copy()
df_old = df_old.sort_values(by='DaysOpen', ascending=False)

# 1.3.3 Preparar gráfico “Distribución by Category”
category_counts = (
    df['Category_norm']
    .value_counts()
    .reset_index()
    .rename(columns={'index': 'Category', 'Category_norm': 'Count'})
)
category_counts.columns = ['Category', 'Count']
category_counts['Percentage'] = (category_counts['Count'] / total_tickets * 100).round(2)

plt.figure(figsize=(8, 6))
top20 = category_counts.head(20)
y_pos = range(len(top20))
plt.barh(y_pos, top20['Count'], align='center')
plt.yticks(y_pos, top20['Category'])
plt.xlabel('Number of Tickets')
plt.title('Distribution by Category (Top 20)')
for i, v in enumerate(top20['Count']):
    plt.text(v + 3, i, str(v), va='center')
plt.gca().invert_yaxis()
plt.tight_layout()

buf = io.BytesIO()
plt.savefig(buf, format='png', dpi=300, bbox_inches='tight')
buf.seek(0)
plt.close()

# ----------------------------------------
# 2. --- Cargar datos de JIRA ---
# ----------------------------------------
jira_path = r'D:\l10 report\Data\Jira Export CSV (all fields) 20250604054311.csv'
if not os.path.isfile(jira_path):
    raise FileNotFoundError(f"No se encontró el archivo de Jira: {jira_path}")

df_jira = pd.read_csv(jira_path, low_memory=False)
df_jira['Created'] = pd.to_datetime(df_jira['Created'], dayfirst=False, utc=True)
df_jira['Created'] = df_jira['Created'].dt.tz_localize(None)
df_jira['DaysOpen'] = (now - df_jira['Created']).dt.days

# 2.1 Filtrar "USA Scaled Tickets" (Labels == 'COLSupport')
df_usa = df_jira[df_jira['Labels'].astype(str).str.contains('COLSupport', na=False)].copy()

# Filtrar solo para los estados "To Do", "In Progress" y "QA"
relevant_statuses = ["To Do", "In Progress", "QA"]
df_usa_filtered = df_usa[df_usa['Status'].isin(relevant_statuses)].copy()

# Actualizar el total de tickets para mostrar solo estos 3 estados
usa_total = len(df_usa_filtered)

# Distribución USA por Status (para el gráfico), solo incluyendo los estados relevantes
usa_status_counts = (
    df_usa_filtered['Status']
    .value_counts()
    .reset_index()
    .rename(columns={'index': 'Status', 'Status': 'Count'})
)
usa_status_counts.columns = ['Status', 'Count']

# Crear gráfico horizontal de barras para USA Scaled Tickets - Status
plt.figure(figsize=(8, 4))
# Usar sort_values para ordenar los estados (To Do, In Progress, QA, etc.)
usa_status_sorted = usa_status_counts.sort_values(by='Count', ascending=True)
y_pos = range(len(usa_status_sorted))

plt.barh(y_pos, usa_status_sorted['Count'], align='center', color='#116693')  # Color azul similar al de la imagen
plt.yticks(y_pos, usa_status_sorted['Status'])
plt.xlabel('Number of Tickets')
plt.title('STATUS')
plt.grid(True, axis='x', linestyle='--', alpha=0.7)

# Añadir el conteo al final de cada barra
for i, v in enumerate(usa_status_sorted['Count']):
    plt.text(v + 0.1, i, str(v), va='center')

plt.tight_layout()

# Guardar gráfico en memoria para insertarlo luego en Word
usa_status_buf = io.BytesIO()
plt.savefig(usa_status_buf, format='png', dpi=300, bbox_inches='tight')
usa_status_buf.seek(0)
plt.close()

# Distribución USA por Priority, filtrando por los estados específicos
relevant_statuses = ["In Progress", "QA", "Release", "To Do"]
df_usa_status_filtered = df_usa[df_usa['Status'].isin(relevant_statuses)].copy()

usa_priority_counts = (
    df_usa_status_filtered['Priority']  # Use the filtered dataframe 
    .value_counts()
    .reset_index()
    .rename(columns={'index': 'Priority', 'Priority': 'Count'})
)
usa_priority_counts.columns = ['Priority', 'Count']

# Crear gráfico horizontal de barras para USA Scaled Tickets - Priority
plt.figure(figsize=(8, 4))
# Ordenar por orden alfabético (Low, Medium, High) o como prefieras
usa_priority_sorted = usa_priority_counts.sort_values(by='Count', ascending=True)
y_pos = range(len(usa_priority_sorted))

plt.barh(y_pos, usa_priority_sorted['Count'], align='center', color='#116693')  # Color azul similar al anterior
plt.yticks(y_pos, usa_priority_sorted['Priority'])
plt.xlabel('Number of Tickets')
plt.title('PRIORITY')
plt.grid(True, axis='x', linestyle='--', alpha=0.7)

# Añadir el conteo al final de cada barra
for i, v in enumerate(usa_priority_sorted['Count']):
    plt.text(v + 0.1, i, str(v), va='center')

plt.tight_layout()

# Guardar gráfico en memoria para insertarlo luego en Word
usa_priority_buf = io.BytesIO()
plt.savefig(usa_priority_buf, format='png', dpi=300, bbox_inches='tight')
usa_priority_buf.seek(0)
plt.close()

usa_avg_days = df_usa.groupby('Priority')['DaysOpen'].mean().round(2)
if not usa_avg_days.empty:
    highest_priority_usa = usa_avg_days.idxmax()
    highest_avg_usa = int(usa_avg_days.max())
else:
    highest_priority_usa = None
    highest_avg_usa = 0

if highest_priority_usa:
    df_usa_top = df_usa[df_usa['Priority'] == highest_priority_usa].copy()
else:
    df_usa_top = pd.DataFrame(columns=df_usa.columns)

# 2.2 “Global Scaled Tickets” (todos los tickets de JIRA)
df_global = df_jira.copy()

# Filtrar por los estados especificados para las tablas Global
relevant_global_statuses = ["In Progress", "QA", "Release", "To Do"]
df_global_filtered = df_global[df_global['Status'].isin(relevant_global_statuses)].copy()

# Actualizar el total para mostrar solo los estados relevantes
global_total = len(df_global_filtered)

# Distribución Global por Status (solo estados relevantes)
global_status_counts = (
    df_global_filtered['Status']
    .value_counts()
    .reset_index()
    .rename(columns={'index': 'Status', 'Status': 'Count'})
)
global_status_counts.columns = ['Status', 'Count']

# Distribución Global por Priority (solo para tickets con estados relevantes)
global_priority_counts = (
    df_global_filtered['Priority']
    .value_counts()
    .reset_index()
    .rename(columns={'index': 'Priority', 'Priority': 'Count'})
)
global_priority_counts.columns = ['Priority', 'Count']

# "Global Tickets with Highest Priority" (filtrado para los estados relevantes)
df_global_top = df_global_filtered[df_global_filtered['Priority'] == 'Highest'].copy()

# ----------------------------------------
# 3. --- Abrir plantilla de Word y actualizar contenido ---
# ----------------------------------------
docx_path = r'D:\l10 report\Reporte base\test de reporte l10.docx'
if not os.path.isfile(docx_path):
    raise FileNotFoundError(f"No se encontró la plantilla de Word: {docx_path}")

doc = Document(docx_path)

# 3.1 Actualizar “Total Tickets:” (soporte) bajo “Support & Tickets Report”
for i, para in enumerate(doc.paragraphs):
    prev_text = doc.paragraphs[i-1].text if i > 0 else ''
    if 'Support & Tickets Report' in prev_text and para.text.strip().startswith('Total Tickets:'):
        para.text = f'Total Tickets: {total_tickets}'
        break

# 3.2 Rellenar TABLA 0 (índice 0): "STATUS  %  # TICKETS" con estilo
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

table0 = doc.tables[0]
# Limpiar la tabla existente, dejando solo la cabecera
for _ in range(len(table0.rows) - 1):
    table0._tbl.remove(table0.rows[1]._tr)

# Estilizar la cabecera con fondo gris oscuro y texto blanco
header_row = table0.rows[0]
for i, cell in enumerate(header_row.cells):
    # Establecer fondo gris oscuro (#404040)
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="404040"/>')
    cell._element.get_or_add_tcPr().append(shading_elm)
    
    # Establecer texto en blanco y centrado
    cell.text = ['STATUS', '%', '# TICKETS'][i]
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in paragraph.runs:
            # Establecer el color como objeto RGBColor en lugar de usar parse_xml
            from docx.shared import RGBColor
            run.font.color.rgb = RGBColor(255, 255, 255)  # Color blanco
            run.font.bold = True
            run.font.size = Pt(11)

# Definir colores para cada status
status_colors = {
    'Done': 'A9D08E',       # Verde para Done
    'In Progress': 'FFD966', # Amarillo para In Progress
    'Scaled': 'BDD7EE',     # Azul claro para Scaled
    "Won't do": 'D9D9D9',   # Gris para Won't do
}

# Insertar filas con estilos que coinciden con la imagen
for _, row_data in status_counts.iterrows():
    row_cells = table0.add_row().cells
    status = row_data['Status']
    
    # Texto de las celdas
    row_cells[0].text = status
    row_cells[1].text = f"{row_data['Percentage']}%"
    row_cells[2].text = str(int(row_data['Count']))
    
    # Aplicar color de fondo según el status
    color = status_colors.get(status, 'FFFFFF')
    for i, cell in enumerate(row_cells):
        # Aplicar sombreado de celda
        shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>')
        cell._element.get_or_add_tcPr().append(shading_elm)
        
        # Centrar contenido para columnas % y # TICKETS
        if i > 0:  # Las columnas 1 y 2 (% y # TICKETS)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Ajustar formato de texto
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(11)

# Asegurar bordes visibles para toda la tabla
# Eliminar la línea que asigna el estilo 'Table Grid'
# table0.style = 'Table Grid'  # Esta línea causa el error

# Aplicar bordes manualmente:
from docx.oxml.shared import OxmlElement

def set_cell_border(cell, **kwargs):
    """
    Set cell border
    Usage:
    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "#000000"},
        bottom={"sz": 12, "val": "single", "color": "#000000"},
        left={"sz": 12, "val": "single", "color": "#000000"},
        right={"sz": 12, "val": "single", "color": "#000000"},
    )
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    for side, attrs in kwargs.items():
        tag = f"w:{side}"
        # Use xpath to find elements - more reliable than direct find
        element_list = tcPr.xpath(f'./{tag}')
        if element_list:
            elm = element_list[0]
        else:
            elm = OxmlElement(tag)
            tcPr.append(elm)
            
        # Fix: When setting attributes, don't include the namespace prefix in the attribute name
        for k, v in attrs.items():
            # Set the attribute without the w: prefix
            elm.set(k, str(v))
                
# Aplicar bordes a todas las celdas
border_attrs = {"sz": 12, "val": "single", "color": "#000000"}
for row in table0.rows:
    for cell in row.cells:
        set_cell_border(
            cell,
            top=border_attrs,
            bottom=border_attrs,
            left=border_attrs,
            right=border_attrs,
        )

# 3.3 Insertar gráfico “Distribution by Category” justo después del párrafo “Distribution by Category:”
for i, para in enumerate(doc.paragraphs):
    if para.text.strip() == 'Distribution by Category:':
        img_para = insert_paragraph_after(para)
        run = img_para.add_run()
        run.add_picture(buf, width=Inches(6))
        break

# 3.4 Rellenar TABLA 1 (índice 1): “Tickets with more than 10 days open”
table1 = doc.tables[1]
for _ in range(len(table1.rows) - 1):
    table1._tbl.remove(table1.rows[1]._tr)
for _, row in df_old.iterrows():
    cells = table1.add_row().cells
    cells[0].text = str(int(row['IDTicket']))
    cells[1].text = str(row['Companyname'])
    cells[2].text = str(row['Reporter'])
    cells[3].text = str(row['Description'])
    cells[4].text = str(int(row['DaysOpen']))

# 3.5 Actualizar narrativa “Tickets > 10 días” (solo IDs)
start_del = None
end_del = None
if not df_old.empty:
    primer_id = str(int(df_old.iloc[0]['IDTicket']))
    for idx, para in enumerate(doc.paragraphs):
        if para.text.strip().startswith(primer_id + ':'):
            start_del = idx
            break
for idx, para in enumerate(doc.paragraphs):
    if 'USA Scaled Tickets' in para.text:
        end_del = idx
        break
if start_del is not None and end_del is not None:
    for _ in range(end_del - start_del):
        doc.paragraphs[start_del]._element.getparent().remove(doc.paragraphs[start_del]._element)

insert_pos = None
for idx, para in enumerate(doc.paragraphs):
    if 'Tickets with more than 10 days open' in para.text:
        insert_pos = idx + 1
        break
if insert_pos is not None:
    for ticket_id in df_old['IDTicket']:
        doc.paragraphs[insert_pos].insert_paragraph_before(f"{int(ticket_id)}:")
        insert_pos += 1

# ----------------------------------------
# 4. --- Sección USA Scaled Tickets ---
# ----------------------------------------
# Encontrar “USA Scaled Tickets” y actualizar la línea siguiente
for i, para in enumerate(doc.paragraphs):
    if 'USA Scaled Tickets' in para.text:
        if i + 1 < len(doc.paragraphs) and doc.paragraphs[i+1].text.strip().startswith('Total tickets:'):
            doc.paragraphs[i+1].text = f'Total tickets: {usa_total}'
        break

# Insertar tabla USA Priority debajo de “Priority” en esa sección
usa_priority_para = None
found_usa_header = False
for para in doc.paragraphs:
    if 'USA Scaled Tickets' in para.text:
        found_usa_header = True
    elif found_usa_header and para.text.strip().startswith('Priority'):
        usa_priority_para = para
        break

if usa_priority_para is not None:
    table_usa = insert_table_after(doc, usa_priority_para, 1, 2)
    hdr = table_usa.rows[0].cells
    hdr[0].text = 'Priority'
    hdr[1].text = '# Tickets'
    for _, row in usa_priority_counts.iterrows():
        r = table_usa.add_row().cells
        r[0].text = str(row['Priority'])
        r[1].text = str(int(row['Count']))

# Reemplazar “Highest average days opened = …”
for para in doc.paragraphs:
    if para.text.strip().startswith('Highest average days opened'):
        para.text = f'Highest average days opened = {highest_avg_usa} days'
        break

# Insertar “USA Tickets with Top Priority” mini-tabla
usa_top_para = None
for para in doc.paragraphs:
    if para.text.strip().startswith('USA Tickets with Top Priority'):
        usa_top_para = para
        break

# Instead of filtering by highest average days, filter for High priority tickets with specific statuses
relevant_statuses_for_top = ["To Do", "In Progress", "QA"]
df_usa_top = df_usa[
    (df_usa['Priority'] == 'High') & 
    (df_usa['Status'].isin(relevant_statuses_for_top))
].copy().sort_values(by='Created', ascending=False)

if usa_top_para is not None:
    # Recreate the table similar to the image provided
    table_usa_top = insert_table_after(doc, usa_top_para, 1, 7)  # Table with 7 columns
    
    # Set column headers like in the image
    hdr = table_usa_top.rows[0].cells
    hdr[0].text = 'Issue key'
    hdr[1].text = 'Issue id'
    hdr[2].text = 'Summary'
    hdr[3].text = 'Priority'
    hdr[4].text = 'Status'
    hdr[5].text = 'Created'
    hdr[6].text = 'Sprint'
    
    # Style the table header
    style_table_like_image(table_usa_top)
    
    # Add data rows
    for _, row in df_usa_top.head(10).iterrows():
        r = table_usa_top.add_row().cells
        r[0].text = str(row['Issue key'])
        r[1].text = str(row.get('Issue id', ''))
        r[2].text = str(row.get('Summary', ''))
        r[3].text = str(row.get('Priority', ''))
        r[4].text = str(row.get('Status', ''))
        
        # Format the Created date to match the image format (MM/DD/YYYY)
        created_date = row.get('Created')
        if created_date is not None and hasattr(created_date, 'strftime'):
            r[5].text = created_date.strftime('%m/%d/%Y')
        else:
            r[5].text = str(created_date)
            
        r[6].text = str(row.get('Sprint', ''))

# Encontrar "USA Scaled Tickets" y actualizar la línea siguiente
for i, para in enumerate(doc.paragraphs):
    if 'USA Scaled Tickets' in para.text:
        # Insertar gráfico de status de USA justo después del párrafo "Total tickets:"
        if i + 1 < len(doc.paragraphs) and doc.paragraphs[i+1].text.strip().startswith('Total tickets:'):
            doc.paragraphs[i+1].text = f'Total tickets: {usa_total}'
            
            # Insertar el gráfico después del total
            img_para = insert_paragraph_after(doc.paragraphs[i+1])
            run = img_para.add_run()
            run.add_picture(usa_status_buf, width=Inches(6))
            
            # Añadir leyenda para el gráfico
            status_caption = insert_paragraph_after(img_para, "Figure: USA Scaled Tickets Status Distribution")
        break

# Insertar el gráfico de Priority en la sección USA Scaled Tickets
for i, para in enumerate(doc.paragraphs):
    if 'USA Scaled Tickets' in para.text:
        # Buscar el párrafo "Priority" en la sección USA
        for j in range(i, len(doc.paragraphs)):
            if doc.paragraphs[j].text.strip() == 'Priority':
                # Insertar el gráfico después del párrafo Priority
                priority_img_para = insert_paragraph_after(doc.paragraphs[j])
                run = priority_img_para.add_run()
                run.add_picture(usa_priority_buf, width=Inches(6))
                
                # Añadir leyenda para el gráfico
                priority_caption = insert_paragraph_after(priority_img_para, "Figure: USA Scaled Tickets Priority Distribution")
                break
        break

# ----------------------------------------
# 5. --- Sección Global Scaled Tickets ---
# ----------------------------------------
# Encontrar “Global Scaled Tickets” y actualizar la línea siguiente
for i, para in enumerate(doc.paragraphs):
    if 'Global Scaled Tickets' in para.text:
        if i + 1 < len(doc.paragraphs) and doc.paragraphs[i+1].text.strip().startswith('Total Tickets:'):
            doc.paragraphs[i+1].text = f'Total Tickets: {global_total}'
        break

# Insertar "Status" + tabla Global Status debajo de "Total Tickets:"
global_status_para = None
for i, para in enumerate(doc.paragraphs):
    if 'Global Scaled Tickets' in para.text:
        if i + 2 < len(doc.paragraphs) and doc.paragraphs[i+1].text.strip().startswith('Status'):
            # si ya existe "Status" en la plantilla, usamos i+1; sino, creamos uno
            global_status_para = doc.paragraphs[i+1]
        else:
            global_status_para = insert_paragraph_after(para, 'Status')
        break

if global_status_para is not None:
    table_global_status = insert_table_after(doc, global_status_para, 1, 2)
    hdr = table_global_status.rows[0].cells
    hdr[0].text = 'Status'
    hdr[1].text = 'Count of Status'
    for _, row in global_status_counts.iterrows():
        r = table_global_status.add_row().cells
        r[0].text = str(row['Status'])
        r[1].text = str(int(row['Count']))
    # Apply the new style to the table
    style_table_like_image(table_global_status)

# Insertar tabla Global Priority debajo de “Priority”
global_priority_para = None
found_global_header = False
for para in doc.paragraphs:
    if 'Global Scaled Tickets' in para.text:
        found_global_header = True
    elif found_global_header and para.text.strip().startswith('Priority'):
        global_priority_para = para
        break

if global_priority_para is not None:
    tbl_glob_prio = insert_table_after(doc, global_priority_para, 1, 2)
    hdr = tbl_glob_prio.rows[0].cells
    hdr[0].text = 'Priority'
    hdr[1].text = 'Count of Priority'
    for _, row in global_priority_counts.iterrows():
        r = tbl_glob_prio.add_row().cells
        r[0].text = str(row['Priority'])
        r[1].text = str(int(row['Count']))
    # Apply the new style to the table
    style_table_like_image(tbl_glob_prio)

# Insertar "Global Tickets with Highest Priority" mini-tabla
glob_top_para = None
for para in doc.paragraphs:
    if para.text.strip().startswith('Global Tickets with Highest Priority'):
        glob_top_para = para
        break

if glob_top_para is not None:
    # Create table with 7 columns to match the image format
    tbl_glob_top = insert_table_after(doc, glob_top_para, 1, 7)
    
    # Set column headers matching the image
    hdr = tbl_glob_top.rows[0].cells
    hdr[0].text = 'Issue key'
    hdr[1].text = 'Issue id'
    hdr[2].text = 'Summary' 
    hdr[3].text = 'Priority'
    hdr[4].text = 'Status'
    hdr[5].text = 'Created'
    hdr[6].text = 'Sprint'
    
    # Apply the blue header styling
    style_table_like_image(tbl_glob_top)
    
    # Add data rows - sort by creation date (newest first)
    sorted_global_top = df_global_top.sort_values(by='Created', ascending=False)
    
    for _, row in sorted_global_top.head(10).iterrows():
        r = tbl_glob_top.add_row().cells
        r[0].text = str(row['Issue key'])
        r[1].text = str(row.get('Issue id', ''))
        r[2].text = str(row.get('Summary', ''))
        r[3].text = str(row.get('Priority', ''))
        r[4].text = str(row.get('Status', ''))
        
        # Format the Created date to match the image format (MM/DD/YYYY)
        created_date = row.get('Created')
        if created_date is not None and hasattr(created_date, 'strftime'):
            r[5].text = created_date.strftime('%m/%d/%Y %H:%M')
        else:
            r[5].text = str(created_date)
            
        r[6].text = str(row.get('Sprint', ''))

# ----------------------------------------
# 6. --- Guardar documento final con fecha en el nombre ---
# ----------------------------------------
fecha = now.strftime('%Y-%m-%d')
hora = now.strftime('%H-%M-%S')
output_dir = r'D:\l10 report\reports'
os.makedirs(output_dir, exist_ok=True)
output_filename = os.path.join(output_dir, f'test_de_reporte_l10_full_{fecha}_time_{hora}.docx')
doc.save(output_filename)

print(f'Reporte generado: {output_filename}')

