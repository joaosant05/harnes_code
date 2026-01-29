import re
import pandas as pd

# ===== Diálogos de arquivo (tkinter) =====
try:
    import tkinter as tk
    from tkinter import filedialog
    TK_OK = True
except Exception:
    TK_OK = False

def select_file():
    if not TK_OK:
        return input("Caminho do arquivo: ").strip()
    root = tk.Tk()
    root.withdraw()
    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo",
        filetypes=[("Planilhas Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
    )
    root.destroy()
    return caminho

def output_path(name="Processed_file.xlsx"):
    if not TK_OK:
        return input(f"Onde salvar o arquivo ({name}): ").strip()
    root = tk.Tk()
    root.withdraw()
    caminho = filedialog.asksaveasfilename(
        title="Onde salvar o Excel de retorno?",
        defaultextension=".xlsx",
        filetypes=[("Planilhas Excel", "*.xlsx")],
        initialfile=name
    )
    root.destroy()
    return caminho

# ===== Variaveis =====
ABA = "SOP - Final"

NOME_COL_WHI = "Wiring Harness Identifier (W)"
NOME_COL_HARNESS = "Harness (PN)"
NOME_COL_VARIANT = "Variant"

ARQ_ENTRADA = select_file()
ARQ_SAIDA = output_path()

# ======= Leitura de TODAS as abas (pra preservar estrutura/valores) =======
xls = pd.ExcelFile(ARQ_ENTRADA)
sheets = {}
for sh in xls.sheet_names:
    sheets[sh] = pd.read_excel(ARQ_ENTRADA, sheet_name=sh, dtype=str)

df = sheets[ABA].copy()

# ======= Lógica WHI =======

# Garante que a coluna WHI exista (se já existir, vamos sobrescrever/preencher)
if NOME_COL_WHI not in df.columns:
    df[NOME_COL_WHI] = ""

# Normaliza vazios
df[NOME_COL_HARNESS] = df[NOME_COL_HARNESS].fillna("")
df[NOME_COL_VARIANT] = df[NOME_COL_VARIANT].fillna("")
df[NOME_COL_WHI] = df[NOME_COL_WHI].fillna("")

# Harness únicos na ordem que aparecem
harness_order = pd.unique(df[NOME_COL_HARNESS])

# Funções auxiliares
v_regex = re.compile(r"^\s*v\s*(\d+)\s*$", re.IGNORECASE)

def variant_key(v: str):
    """
    Ordena v1..vx por número; 'all' vai por último.
    Qualquer outra coisa também vai depois (mas mantendo comportamento previsível).
    """
    if v is None:
        return (2, 10**9)
    vv = v.strip().lower()
    if vv == "all":
        return (1, 10**9)
    m = v_regex.match(vv)
    if m:
        return (0, int(m.group(1)))
    return (2, 10**9)

w_counter = 1  # contador global

for h in harness_order:
    if h == "":
        continue

    mask_h = (df[NOME_COL_HARNESS] == h)
    variants_in_h = pd.unique(df.loc[mask_h, NOME_COL_VARIANT])

    # Separa v1..vx e all
    v_list = []
    has_all = False
    for v in variants_in_h:
        vv = (v or "").strip()
        if vv.lower() == "all":
            has_all = True
        else:
            # só considera vN como variante numerada
            if v_regex.match(vv.lower()):
                v_list.append(vv)
            else:
                # se aparecer algo fora do padrão, você pode ignorar ou tratar.
                # aqui: ignora silenciosamente (sem fallback/validação, conforme seu obs3).
                pass

    # Ordena v1..vx por número
    v_list = sorted(set(v_list), key=lambda x: variant_key(x))

    # Mapa de variantes -> W para este harness
    whi_parts = []

    # Aplica W1, W2, ... para v1..vx (incremental global)
    for v in v_list:
        w_label = f"W{w_counter}"
        w_counter += 1
        mask = mask_h & (df[NOME_COL_VARIANT].str.strip().str.lower() == v.strip().lower())
        df.loc[mask, NOME_COL_WHI] = w_label
        whi_parts.append(w_label)

    # Aplica ALL como junção W1/W2/...
    if has_all:
        joined = "/".join(whi_parts)
        mask_all = mask_h & (df[NOME_COL_VARIANT].str.strip().str.lower() == "all")
        df.loc[mask_all, NOME_COL_WHI] = joined

# Atualiza a aba final no dicionário
sheets[ABA] = df

# ======= Escrita: mesma estrutura/abas (valores) =======
with pd.ExcelWriter(ARQ_SAIDA, engine="openpyxl") as writer:
    for sh_name, sh_df in sheets.items():
        sh_df.to_excel(writer, sheet_name=sh_name, index=False)

print(f"Arquivo salvo em: {ARQ_SAIDA}")
