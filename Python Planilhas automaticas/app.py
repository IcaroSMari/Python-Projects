import pandas as pd
import sys
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime

# Força UTF-8 no print (Windows)
sys.stdout.reconfigure(encoding="utf-8")

# Data atual para nome do arquivo
data_hoje = datetime.today().strftime("%Y-%m-%d")

# Lendo do Excel
df = pd.read_excel("courseid_685_participants.xlsx", sheet_name="Dados")
df.rename(columns=lambda x: x.strip(), inplace=True)

# 🔹 Filtrar apenas estudantes
df_estudantes = df[df["Papel"].str.upper().str.contains("ESTUDANTE")].copy()

# Conta alunos que nunca acessaram
nunca_acessaram = df_estudantes["Ultimo Acesso"].isin(["NUNCA", "Nunca", "nunca", None, ""]).sum()

# Conta alunos que acessaram ao menos uma vez
acessaram = len(df_estudantes) - nunca_acessaram

# Total de alunos
total = len(df_estudantes)

# Percentuais
perc_acessaram = acessaram / total * 100 if total > 0 else 0
perc_nunca = nunca_acessaram / total * 100 if total > 0 else 0

print("Relatório de Acessos - Apenas Estudantes")
print(f"Total de estudantes: {total}")
print(f"Acessaram ao menos uma vez: {acessaram} ({perc_acessaram:.2f}%)")
print(f"Nunca acessaram: {nunca_acessaram} ({perc_nunca:.2f}%)")

# -------------------------------
# Gerar gráfico de pizza
# -------------------------------
labels = ["Acessaram", "Nunca acessaram"]
sizes = [acessaram, nunca_acessaram]
colors_graph = ["#4CAF50", "#F44336"]

nome_grafico = f"grafico_acessos_{data_hoje}.png"

plt.figure(figsize=(5, 5))
plt.pie(sizes, labels=labels, autopct="%.1f%%", colors=colors_graph, startangle=90)
plt.title("Acessos dos Estudantes")
plt.savefig(nome_grafico)
plt.close()

# -------------------------------
# Preparar dados da tabela
# -------------------------------
df_estudantes["Ultimo Acesso"].fillna("Nunca", inplace=True)

# Cabeçalho + dados
tabela_dados = [["Nome", "Último Acesso"]]
for _, row in df_estudantes.iterrows():
    tabela_dados.append([row["Nome"], row["Ultimo Acesso"]])

# -------------------------------
# Criar PDF com gráfico e tabela
# -------------------------------
nome_pdf = f"relatorio_acessos_{data_hoje}.pdf"
doc = SimpleDocTemplate(nome_pdf, pagesize=A4)
elements = []
styles = getSampleStyleSheet()

# Título
elements.append(Paragraph("Relatório de Acessos - Estudantes", styles["Title"]))
elements.append(Spacer(1, 20))

# Adiciona gráfico
elements.append(Image(nome_grafico, width=300, height=300))
elements.append(Spacer(1, 20))

# Criar tabela
table = Table(tabela_dados, colWidths=[250, 150])
style = TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ("FONTSIZE", (0, 0), (-1, 0), 12),
    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
])

# Destacar em vermelho os que nunca acessaram
for i, row in enumerate(tabela_dados[1:], start=1):  # pula cabeçalho
    if str(row[1]).strip().lower() == "nunca":
        style.add("TEXTCOLOR", (0, i), (-1, i), colors.red)

table.setStyle(style)
elements.append(table)

# Gerar PDF
doc.build(elements)
print(f"Relatório salvo em {nome_pdf}")
