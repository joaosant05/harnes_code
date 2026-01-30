import re
import pandas as pd

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

def make_unique_columns(cols):
    counts = {}
    out = []
    for c in cols:
        c = "" if c is None else str(c).strip()
        if c not in counts:
            counts[c] = 1
            out.append(c)
        else:
            counts[c] += 1
            out.append(f"{c}__{counts[c]}")
    return out

def read_sheet_with_dynamic_header(path, sheet_name, required_cols, max_scan_rows=80, dtype=str):
    print(f"[1/6] Lendo aba '{sheet_name}' (header dinâmico)...")
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=dtype)

    def norm(x):
        if x is None:
            return ""
        s = str(x)
        s = re.sub(r"\s+", " ", s).strip()
        return s

    required_norm = [norm(c) for c in required_cols]

    header_row = None
    scan = min(max_scan_rows, len(raw))
    for i in range(scan):
        row_vals = [norm(v) for v in raw.iloc[i].tolist()]
        row_set = set(row_vals)
        if all(c in row_set for c in required_norm):
            header_row = i
            break

    if header_row is None:
        raise ValueError(
            f"Não encontrei o cabeçalho nas primeiras {max_scan_rows} linhas da aba '{sheet_name}'. "
            f"Precisava conter as colunas: {required_cols}"
        )

    headers = [norm(v) for v in raw.iloc[header_row].tolist()]
    headers = make_unique_columns(headers)

    df = raw.iloc[header_row + 1:].copy()
    df.columns = headers
    df.reset_index(drop=True, inplace=True)
    print(f"[2/6] Header encontrado na linha {header_row+1}. Linhas: {len(df):,} | Colunas: {len(df.columns)}")
    return df, header_row

def is_filled_series(s: pd.Series) -> pd.Series:
    s = s.astype(str).fillna("").str.strip()
    return (s != "") & (s.str.lower() != "nan")

ARQ_ENTRADA = select_file()
if not ARQ_ENTRADA:
    raise SystemExit("Nenhum arquivo selecionado.")

xls = pd.ExcelFile(ARQ_ENTRADA)
print("\nAbas disponíveis:")
for sh in xls.sheet_names:
    print(f" - {sh}")

ABA = input("\nInforme o nome da planilha que deseja analisar: ").strip()
if ABA not in xls.sheet_names:
    raise ValueError(f"A aba '{ABA}' não existe no arquivo.")

HARNESS_COL = input("Informe o nome da coluna de chicotes (HARNESS): ").strip()

NOME_COL_WHI = "Wiring Harness Identifier (W)"
NOME_COL_MIDDLE = "Logic Connector Identifier"
NOME_COL_LEFT = "lado esquerdo"
NOME_COL_RIGHT = "lado direito"

ARQ_SAIDA = output_path()
if not ARQ_SAIDA:
    raise SystemExit("Caminho de saída não informado.")

print("\n[0/6] Lendo todas as abas (pode demorar em arquivos grandes)...")
sheets = {}
for sh in xls.sheet_names:
    print(f"  - Lendo aba: {sh}")
    sheets[sh] = pd.read_excel(ARQ_ENTRADA, sheet_name=sh, dtype=str)

required = [HARNESS_COL, NOME_COL_MIDDLE, NOME_COL_RIGHT, NOME_COL_LEFT]
df, header_row = read_sheet_with_dynamic_header(ARQ_ENTRADA, ABA, required_cols=required, dtype=str)

if HARNESS_COL not in df.columns:
    raise ValueError(f"Não achei a coluna '{HARNESS_COL}' no cabeçalho. Colunas: {list(df.columns)}")

missing = [c for c in [HARNESS_COL, NOME_COL_MIDDLE, NOME_COL_RIGHT, NOME_COL_LEFT] if c not in df.columns]
if missing:
    raise ValueError(f"Colunas não encontradas: {missing}\nColunas disponíveis: {list(df.columns)}")

if NOME_COL_WHI not in df.columns:
    df[NOME_COL_WHI] = ""

print("[3/6] Preparando códigos (um por código, ordem de cima pra baixo)...")
harness = df[HARNESS_COL].astype(str).fillna("").str.strip()
codes_in_order = pd.unique(harness[harness != ""])
print(f"  - Códigos únicos: {len(codes_in_order)}")

groups = df.groupby(harness, sort=False).groups
priority_cols = [NOME_COL_MIDDLE, NOME_COL_RIGHT, NOME_COL_LEFT]

print("[4/6] Aplicando regra de prioridade e atribuindo W...")
w_counter = 1
processed = 0
assigned_count = 0

for code in codes_in_order:
    idx = groups.get(code)
    if idx is None or len(idx) == 0:
        continue

    assigned = False
    for col in priority_cols:
        if is_filled_series(df.loc[idx, col]).any():
            df.loc[idx, NOME_COL_WHI] = f"W{w_counter}"
            w_counter += 1
            assigned = True
            assigned_count += 1
            break

    processed += 1
    if processed % 200 == 0:
        print(f"  ...{processed}/{len(codes_in_order)} códigos processados (W atribuídos: {assigned_count})")

print(f"[5/6] Concluído. Códigos processados: {processed} | W atribuídos: {assigned_count}")

sheets[ABA] = df

print("[6/6] Salvando arquivo de saída...")
with pd.ExcelWriter(ARQ_SAIDA, engine="openpyxl") as writer:
    for sh, sdf in sheets.items():
        sdf.to_excel(writer, sheet_name=sh, index=False)

print(f"✅ Arquivo gerado em: {ARQ_SAIDA}")
