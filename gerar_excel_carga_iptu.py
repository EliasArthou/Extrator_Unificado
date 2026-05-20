"""
gerar_excel_carga_iptu.py — Gera planilha Excel pra carga manual em [Codigos IPTUs].

Varre uma pasta de PDFs de boletos IPTU (formato cod_iptu.pdf), extrai os campos
de cada boleto usando aux.extrairtextopdf(..., 'Boletos') — a MESMA função que
o pipeline usa — e consolida tudo num XLSX pronto pra importação manual.

Uso:
    python gerar_excel_carga_iptu.py                  # default: Downloads/IPTU/
    python gerar_excel_carga_iptu.py -p outra/pasta   # pasta diferente
    python gerar_excel_carga_iptu.py -o saida.xlsx    # nome do arquivo de saída

Colunas geradas (mesma ordem e nomes da tabela [Codigos IPTUs]):
    CODIGO, Cod IPTUs, Competencia, DataPgto, TpoPagto, Arquivo,
    NomeBoleto, Barras, Nr Guia, Valor, Parcela

Observações:
- Valor sai como float (ex: 1087.40). Quando colar no banco, o Access converte.
- Parcela sai como int.
- Cada PDF gera MÚLTIPLAS linhas (uma por parcela).
- TpoPagto fixo em 'PARCELADO' (mesma convenção do Biptu_hibrido.py:352).
- Arquivos que falharem no parse vão pra aba 'Falhas' do Excel.
"""

import argparse
import os
import sys
from datetime import datetime

import pandas as pd
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TimeElapsedColumn

import auxiliares as aux

console = Console()


def main():
    parser = argparse.ArgumentParser(description="Gera Excel pra carga em [Codigos IPTUs]")
    parser.add_argument("-p", "--pasta", default="Downloads/IPTU",
                        help="Pasta com os PDFs (default: Downloads/IPTU)")
    parser.add_argument("-o", "--saida", default="",
                        help="Arquivo de saida .xlsx (default: Logs/Carga_IPTU_<timestamp>.xlsx)")
    args = parser.parse_args()

    pasta = os.path.abspath(args.pasta)
    if not os.path.isdir(pasta):
        console.print(f"[red][ERRO][/] Pasta nao encontrada: {pasta}")
        sys.exit(1)

    pdfs = sorted([
        os.path.join(pasta, f)
        for f in os.listdir(pasta)
        if f.lower().endswith('.pdf')
    ])
    if not pdfs:
        console.print(f"[yellow][AVISO][/] Nenhum PDF em {pasta}")
        sys.exit(0)

    console.print(f"[cyan][INFO][/] {len(pdfs)} PDF(s) em {pasta}")

    saida = args.saida
    if not saida:
        ts = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
        pasta_log = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Logs")
        os.makedirs(pasta_log, exist_ok=True)
        saida = os.path.join(pasta_log, f"Carga_IPTU_{ts}.xlsx")
    saida = os.path.abspath(saida)

    dfs = []
    falhas = []

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TextColumn("{task.completed}/{task.total}"),
        TimeElapsedColumn(),
        console=console,
    ) as progress:
        task = progress.add_task("Extraindo PDFs...", total=len(pdfs))

        for caminho in pdfs:
            nome = os.path.basename(caminho)
            base = os.path.splitext(nome)[0]

            if '_' not in base:
                falhas.append({'Arquivo': nome, 'Erro': "Nome fora do padrao cod_iptu.pdf"})
                progress.advance(task)
                continue
            cod, _ = base.split('_', 1)

            try:
                df = aux.extrairtextopdf(caminho, 'Boletos')
            except Exception as e:
                falhas.append({'Arquivo': nome, 'Erro': f"Excecao: {e}"})
                progress.advance(task)
                continue

            if df is None or len(df) == 0:
                falhas.append({'Arquivo': nome, 'Erro': "extrairtextopdf retornou vazio"})
                progress.advance(task)
                continue

            # Tratamento equivalente ao Biptu_hibrido.py:346-357 + extracao.py:632-640:
            #  - remove aspas simples de todas as strings
            #  - converte 'Valor Total' pra float (1.087,40 -> 1087.40)
            #  - converte 'Parcela' pra int
            #  - insere CODIGO (do nome), TpoPagto='PARCELADO', Arquivo
            try:
                df = df.map(
                    lambda x: x.replace("'", "") if isinstance(x, str) else x
                )
                df['Valor Total'] = (df['Valor Total']
                                     .str.replace('.', '', regex=False)
                                     .str.replace(',', '.', regex=False)
                                     .astype('float'))
                df['Parcela'] = df['Parcela'].astype(int)
            except Exception as e:
                falhas.append({'Arquivo': nome, 'Erro': f"Falha na conversao de tipos: {e}"})
                progress.advance(task)
                continue

            df.insert(0, 'Codigo', cod)
            df.insert(4, 'TpoPagto', 'PARCELADO')
            df.insert(5, 'Arquivo', nome)

            # Renomeia pros nomes oficiais da tabela [Codigos IPTUs]
            df = df.rename(columns={
                'Codigo': 'CODIGO',
                'Inscricao': 'Cod IPTUs',
                'Vencimentos': 'DataPgto',
                'Contribuinte': 'NomeBoleto',
                'Codigo de Barras': 'Barras',
                'Guia': 'Nr Guia',
                'Valor Total': 'Valor',
            })

            dfs.append(df)
            progress.advance(task)

    if not dfs:
        console.print("[red][ERRO][/] Nenhum PDF extraido com sucesso. Veja a aba Falhas.")

    consolidado = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame(
        columns=['CODIGO', 'Cod IPTUs', 'Competencia', 'DataPgto', 'TpoPagto',
                 'Arquivo', 'NomeBoleto', 'Barras', 'Nr Guia', 'Valor', 'Parcela']
    )

    with pd.ExcelWriter(saida, engine='openpyxl') as writer:
        consolidado.to_excel(writer, sheet_name='Carga', index=False)
        if falhas:
            pd.DataFrame(falhas).to_excel(writer, sheet_name='Falhas', index=False)

    console.print(f"\n[bold]{'='*60}[/]")
    console.print(f"[bold green]RESUMO[/]")
    console.print(f"[bold]{'='*60}[/]")
    console.print(f"  PDFs processados: {len(pdfs)}")
    console.print(f"  [green]Sucesso:[/]          {len(dfs)}")
    console.print(f"  [red]Falhas:[/]           {len(falhas)}")
    console.print(f"  Linhas geradas:   {len(consolidado)}  (1 por parcela)")
    console.print(f"  [cyan]Arquivo:[/]          {saida}")

    if falhas:
        console.print(f"\n[bold red]Falhas:[/]")
        for f in falhas[:10]:
            console.print(f"  {f['Arquivo']}  |  {f['Erro']}")
        if len(falhas) > 10:
            console.print(f"  ... + {len(falhas) - 10} mais (ver aba Falhas)")


if __name__ == "__main__":
    main()
