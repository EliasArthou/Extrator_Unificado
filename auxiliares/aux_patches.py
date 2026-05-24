"""
aux_patches.py — Monkey-patches para o modulo `auxiliares`.

Este arquivo adiciona ao modulo `auxiliares` funcoes que sao chamadas pelo
projeto mas que nunca foram implementadas. O proposito e nao tocar no
arquivo `auxiliares.py` original (que tem caracteristicas de encoding
sensiveis a edicao automatica).

USO:
    Apos `import auxiliares as aux`, adicionar `import aux_patches` para
    aplicar os patches:

        from auxiliares import utils as aux
        from auxiliares import aux_patches  # registra patches no modulo auxiliares

    Depois disso, `aux.adicionarcabecalhopdf_topo_adaptativo(...)` funciona.

PATCHES ATUAIS:
    - adicionarcabecalhopdf_topo_adaptativo: wrapper de compatibilidade.
      O codigo do projeto chama essa funcao em ~20 lugares, mas ela nunca
      foi implementada. Sem este patch, todas as chamadas falham com
      AttributeError (engolido por try/except), resultando em PDFs sem
      cabecalho. Este patch delega para a funcao base
      `adicionarcabecalhopdf()` que esta implementada e funciona.
"""

from auxiliares import utils as _aux


def adicionarcabecalhopdf_topo_adaptativo(pdf_entrada, pdf_saida, codigo,
                                          senha_base=None, margem=5):
    """
    Wrapper de compatibilidade — DEPRECATED.

    Mantido apenas para casos em que `senha_base` (ex: CPF) e' fornecido
    explicitamente. A versao base em auxiliares.adicionarcabecalhopdf
    apaga o arquivo de entrada (os.remove na linha 982), o que QUEBRA
    quando pdf_entrada == pdf_saida (caso de extracao_api.py:2262).

    Por isso o monkey-patch foi DESATIVADO. A versao em
    auxiliares.py:992 ja' implementa esse fluxo corretamente com
    tempfile + shutil.move e funciona com mesma entrada/saida.
    """
    if senha_base:
        try:
            import pikepdf
            with pikepdf.open(pdf_entrada, password=str(senha_base)) as pdf:
                if pdf.is_encrypted:
                    pdf.save(pdf_entrada)
        except Exception:
            pass

    return _aux.adicionarcabecalhopdf(pdf_entrada, pdf_saida, codigo)


# Monkey-patch DESATIVADO (2026-05-21):
#   A versao `_topo_adaptativo` ja' existe em auxiliares.py:992 com logica
#   correta (tempfile + shutil.move, sem os.remove do entrada). Sobrescrever
#   por este wrapper estava jogando pra `adicionarcabecalhopdf` (versao
#   antiga na linha 934) que apaga o PDF quando entrada==saida, resultando
#   em PDFs sem cabecalho silenciosamente (excecao engolida pelo try/except
#   em extracao_api.py:2263).
#
# Se algum codigo precisar do wrapper com `senha_base`, chame
# `aux_patches.adicionarcabecalhopdf_topo_adaptativo` diretamente em vez
# de via `aux.adicionarcabecalhopdf_topo_adaptativo`.
#
# Linha original (mantida em comentario para referencia):
# _aux.adicionarcabecalhopdf_topo_adaptativo = adicionarcabecalhopdf_topo_adaptativo
