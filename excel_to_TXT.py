#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
PortabilidadeExcelTXT — Excel → TXT (interativo)

Versão 1.0.3
- Mensagem de arquivo incompatível simplificada (sem detalhes técnicos).
- Evita repetição do "Clique Enter para fechar esta tela...".
- Mostra apenas o nome do arquivo nas mensagens.
- Pausas: pasta vazia, erros, conclusão.

Funcionalidades principais:
- Detecta automaticamente DATA (DD/MM/AAAA) e duas colunas monetárias ADJACENTES.
- Converte DATA → AAAAMMDD; VALOR 1/2 → centavos inteiros.
- Remove possível linha de somatório no rodapé (DATA vazia + totalizador).
- Suporte a arquivos protegidos por senha (msoffcrypto-tool + getpass).
"""


import argparse
import re
import sys
import io
import getpass
from pathlib import Path
import pandas as pd

__app_name__ = "PortabilidadeExcelTXT"
__version__  = "1.0.3"

# ================== Base dir (funciona em .py e .exe) ==================
if getattr(sys, "frozen", False):  # rodando como executável
    BASE_DIR = Path(sys.executable).parent
else:  # rodando como script
    BASE_DIR = Path(__file__).parent

PASTA_ENTRADA = BASE_DIR / "entradas"
PASTA_SAIDA   = BASE_DIR / "saidas"

def ensure_dirs():
    PASTA_ENTRADA.mkdir(parents=True, exist_ok=True)
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)

# ================== Helpers de UX ==================
def wait_to_close(msg_suffix: str = "Clique Enter para fechar esta tela..."):
    """Exibe um prompt de pausa único, sem duplicar mensagens."""
    try:
        input(f"\n{msg_suffix}")
    except EOFError:
        pass  # Ambientes sem stdin

def pause_and_exit(message: str, exit_code: int = 0):
    """Imprime a mensagem e aguarda Enter antes de sair."""
    print(message)
    wait_to_close()
    sys.exit(exit_code)

# ================== Config TXT fixo ==================
STRING_FIXA   = "03654036541584001"
TAMANHO_LINHA = 1000
# 02 (2) + seq(6) + STRING_FIXA(17) + DATA(8) + V1(15) + V2(15) = 63
ESPACO_EXTRA  = " " * (TAMANHO_LINHA - 63)
SEQ_INICIAL   = 3  # sequencial inicial

# ================== Regras de detecção Excel ==================
DATE_RX = re.compile(r"^\s*(\d{2})/(\d{2})/(\d{4})\s*$")
MONEY_RXS = [
    re.compile(r"^\s*R\$\s*\d{1,3}(?:\.\d{3})*,\d{2}\s*$"),
    re.compile(r"^\s*R\$\s*\d+,\d{2}\s*$"),
    re.compile(r"^\s*\d{1,3}(?:\.\d{3})*,\d{2}\s*$"),
    re.compile(r"^\s*\d+,\d{2}\s*$"),
]

# ================== Utilidades de parsing ==================
def is_date_like(val) -> bool:
    if pd.isna(val): return False
    if isinstance(val, pd.Timestamp): return True
    return bool(DATE_RX.match(str(val).strip()))

def parse_date_to_yyyymmdd(val) -> str:
    if pd.isna(val): return ""
    if isinstance(val, pd.Timestamp): return val.strftime("%Y%m%d")
    s = str(val).strip()
    m = DATE_RX.match(s)
    if m:
        d, mo, y = m.groups()
        return f"{y}{mo}{d}"
    try:
        return pd.to_datetime(s, dayfirst=True, errors="raise").strftime("%Y%m%d")
    except Exception:
        return ""

def is_money_like(val) -> bool:
    if pd.isna(val): return False
    if isinstance(val, (int, float)): return True
    return any(rx.match(str(val).strip()) for rx in MONEY_RXS)

def parse_money_to_centavos_int(val) -> int:
    if pd.isna(val): return 0
    if isinstance(val, (int, float)): return int(round(float(val) * 100))
    s = str(val).replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
    try:
        return int(round(float(s) * 100))
    except Exception:
        return 0

def detect_columns(df: pd.DataFrame):
    date_scores = {c: df[c].apply(is_date_like).mean() for c in df.columns}
    money_scores = {c: df[c].apply(is_money_like).mean() for c in df.columns}
    date_col = max(date_scores, key=date_scores.get)
    best_pair, best_score = (None, None), -1
    cols = list(df.columns)
    for i in range(len(cols)-1):
        c1, c2 = cols[i], cols[i+1]
        score = money_scores[c1] + money_scores[c2]
        if score > best_score:
            best_pair, best_score = (c1, c2), score
    return date_col, best_pair[0], best_pair[1]

def standardize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Mantém somente DATA (AAAAMMDD), VALOR 1/2 (centavos int)."""
    date_col, m1, m2 = detect_columns(df)
    if not date_col or not m1 or not m2:
        raise ValueError("Não foi possível detectar DATA, VALOR 1 e VALOR 2.")
    out = pd.DataFrame()
    out["DATA"]    = df[date_col].apply(parse_date_to_yyyymmdd)
    out["VALOR 1"] = df[m1].apply(parse_money_to_centavos_int)
    out["VALOR 2"] = df[m2].apply(parse_money_to_centavos_int)
    return out

def remove_possible_total_footer(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remove possível linha de somatório no final:
    - DATA vazia ("")
    - VALOR 1/2 ~= soma anterior (±1 centavo) OU muito maiores (>=5x mediana)
    """
    if df.empty or "DATA" not in df or "VALOR 1" not in df or "VALOR 2" not in df:
        return df

    blank_idx = df.index[df["DATA"].eq("")].tolist()
    if not blank_idx:
        return df

    idx = blank_idx[-1]
    row = df.loc[idx]
    prev = df.drop(index=idx)
    if prev.empty:
        return df

    try:
        v1 = int(row["VALOR 1"]); v2 = int(row["VALOR 2"])
    except Exception:
        return df

    sum1 = int(prev["VALOR 1"].sum())
    sum2 = int(prev["VALOR 2"].sum())
    median1 = prev["VALOR 1"].median() if not prev["VALOR 1"].empty else 0
    median2 = prev["VALOR 2"].median() if not prev["VALOR 2"].empty else 0

    cond_equal = (abs(v1 - sum1) <= 1) or (abs(v2 - sum2) <= 1)
    cond_large = (median1 > 0 and v1 >= 5 * median1) or (median2 > 0 and v2 >= 5 * median2)

    if cond_equal or cond_large:
        return prev.reset_index(drop=True)
    return df

# ================== Mensagens de erro amigáveis (simplificadas) ==================
def incompatible_file_message(path: Path) -> str:
    """Mensagem curta e clara para arquivo incompatível (sem detalhes técnicos)."""
    return (
        f"\n❌ Arquivo incompatível: {path.name}\n"
        "Abra no Excel e salve novamente como .xlsx (sem senha)."
    )

# ================== Leitura do Excel (com senha) ==================
def load_df_excel(path: Path) -> pd.DataFrame:
    """
    Lê .xlsx (openpyxl) e .xls (xlrd).
    Se estiver protegido por senha, pede a senha (getpass) e tenta abrir com msoffcrypto-tool.
    Se falhar, orienta a salvar o arquivo SEM SENHA antes de continuar (com pausa).
    """
    suf = path.suffix.lower()

    # 1) Tenta ler normalmente
    try:
        if suf == ".xlsx":
            return pd.read_excel(path, engine="openpyxl")
        if suf == ".xls":
            return pd.read_excel(path, engine="xlrd")
        # Formato não suportado
        pause_and_exit(incompatible_file_message(path), exit_code=2)
    except Exception as e:
        msg = str(e).lower()

        protected_hints = [
            "encrypted", "password", "protected", "is encrypted", "bad password",
            "workbook is password protected", "file is password-protected",
        ]
        is_protected = any(h in msg for h in protected_hints)

        if not is_protected:
            # Qualquer erro de leitura (incluindo .xls não suportado) -> mensagem simples
            pause_and_exit(incompatible_file_message(path), exit_code=2)

        # 2) Parece arquivo protegido → tenta descriptografar com msoffcrypto-tool
        try:
            import msoffcrypto
        except ImportError:
            pause_and_exit(
                "\n🔒 O arquivo parece estar PROTEGIDO POR SENHA.\n"
                "Abra no Excel e SALVE SEM SENHA antes de converter.",
                exit_code=2,
            )

        # Pede a senha no terminal (até 3 tentativas)
        for tentativa in range(3):
            pwd = getpass.getpass(prompt=f"Senha do arquivo '{path.name}' (tentativa {tentativa+1}/3): ")
            if pwd is None:
                pwd = ""
            try:
                with open(path, "rb") as f:
                    office = msoffcrypto.OfficeFile(f)
                    office.load_key(password=pwd)
                    decrypted = io.BytesIO()
                    office.decrypt(decrypted)
                    decrypted.seek(0)
                if suf == ".xlsx":
                    return pd.read_excel(decrypted, engine="openpyxl")
                else:
                    return pd.read_excel(decrypted, engine="xlrd")
            except Exception:
                print("❌ Senha incorreta ou falha ao descriptografar.")

        # 3 tentativas sem sucesso
        pause_and_exit(
            "\n❌ Não foi possível abrir o arquivo protegido.\n"
            "Verifique a senha e tente novamente, ou SALVE o Excel SEM SENHA e rode de novo.",
            exit_code=2,
        )

# ================== Seleção interativa ==================
def escolher_excel() -> Path:
    ensure_dirs()
    arquivos = sorted(list(PASTA_ENTRADA.glob("*.xls")) + list(PASTA_ENTRADA.glob("*.xlsx")))
    # ignora temporários e arquivos já padronizados
    arquivos = [a for a in arquivos if not a.name.startswith("~$") and "_padronizado" not in a.stem.lower()]
    if not arquivos:
        pause_and_exit(
            "❌ Nenhum arquivo .xls/.xlsx encontrado na pasta 'entradas/'.\n"
            "Coloque o arquivo na pasta 'entradas' e rode novamente.",
            exit_code=1,
        )

    print(f"\n{__app_name__} v{__version__}")
    print("📂 Excel disponíveis em 'entradas/':")
    for i, arq in enumerate(arquivos, 1):
        print(f" {i}. {arq.name}")
    while True:
        try:
            escolha = int(input("\nDigite o número do arquivo desejado: "))
            if 1 <= escolha <= len(arquivos):
                return arquivos[escolha - 1]
            else:
                print("Número fora da faixa. Tente novamente.")
        except ValueError:
            print("Entrada inválida. Digite apenas o número.")

# ================== Excel -> TXT ==================
def dataframe_to_fixed_txt(df: pd.DataFrame, out_path: Path, seq_inicial: int = SEQ_INICIAL):
    """
    Gera TXT no layout fixo. Usa VALOR 1/2 em CENTAVOS (já inteiros) e DATA AAAAMMDD.
    """
    out_path.parent.mkdir(parents=True, exist_ok=True)
    linhas = []
    seq = seq_inicial
    for _, row in df.iterrows():
        data = str(row["DATA"])[:8].rjust(8, "0") if row["DATA"] else "00000000"
        v1 = int(row["VALOR 1"])
        v2 = int(row["VALOR 2"])
        linha = f"02{str(seq).zfill(6)}{STRING_FIXA}{data}{v1:015d}{v2:015d}{ESPACO_EXTRA}"
        if len(linha) != TAMANHO_LINHA:
            print(f"⚠ Linha {seq} com tamanho {len(linha)} (esperado {TAMANHO_LINHA}).")
        linhas.append(linha)
        seq += 1
    # CRLF para compatibilidade
    out_path.write_text("\r\n".join(linhas), encoding="utf-8", newline="\r\n")
    return len(linhas)

def run_excel_para_txt_interativo(overwrite: bool = False):
    alvo = escolher_excel()
    print(f"\n▶ Excel→TXT: {alvo.name}")
    try:
        df = load_df_excel(alvo)
        std = standardize_dataframe(df)

        # 1) Ordena por DATA crescente (vazios ao fim)
        sort_key = pd.to_numeric(std["DATA"].replace("", pd.NA), errors="coerce")
        std = std.assign(_key=sort_key).sort_values(by="_key", na_position="last").drop(columns="_key").reset_index(drop=True)

        # 2) Remove possível somatório no rodapé
        std = remove_possible_total_footer(std)

        # 3) Gera TXT fixo
        out = PASTA_SAIDA / (alvo.stem + ".txt")
        if out.exists() and not overwrite:
            print(f"⏭ Saída já existe, não sobrescrevi (use --overwrite): {out}")
            wait_to_close()
            return
        n = dataframe_to_fixed_txt(std, out)
        print(f"✅ TXT gerado: {out}  ({n} linhas)")
        wait_to_close("Processo concluído. Clique Enter para fechar esta tela...")
    except SystemExit:
        # Mensagem e pausa já foram tratadas por pause_and_exit
        pass
    except PermissionError as e:
        pause_and_exit(f"🔒 Arquivo bloqueado/aberto: {alvo.name} — {e}", exit_code=2)
    except Exception:
        pause_and_exit(incompatible_file_message(alvo), exit_code=2)

# ================== CLI ==================
def main():
    ap = argparse.ArgumentParser(description=f"{__app_name__} v{__version__} — Excel → TXT (interativo).")
    ap.add_argument("--overwrite", action="store_true", help="Sobrescrever o TXT se já existir")
    args = ap.parse_args()
    ensure_dirs()
    run_excel_para_txt_interativo(overwrite=args.overwrite)

if __name__ == "__main__":
    main()
    