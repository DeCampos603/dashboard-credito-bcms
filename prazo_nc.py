"""Prazo de uma Nota de Crédito lido do próprio texto da NC (campo "NC - Descrição" do SIAFI).

A planilha do Tesouro Gerencial não tem coluna de prazo: ele só existe no texto livre
que o órgão concedente escreve na NC. Formatos encontrados na base de 2026:

    PRAZO PARA EMPENHO: 06/11/2026      EMPENHAR ATÉ 30/MAR/26     EMPH ATE 15OUT26
    EMPENHO ATÉ 30SET2026               EMPENHO ATE 30 OUT 26      EMPH ATE 15 MAIO 26
    EMPH ATE 48H                        EMPH EM ATE 5 DIAS UTEIS   PRAZO DE EMPENHO: IMEDIATO
    LIQUIDAR ATÉ 30 SET 26              RECOLHER SALDOS ATE 01 NOV 26

Regra: nunca inventar prazo. Se o texto não traz prazo, devolve None e a tela mostra
"Não informado na NC".
"""
import datetime
import re
import unicodedata

_MESES = {
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12,
}

# dia + mês (sigla, com ou sem separador; número, só com / - .) + ano opcional.
# Mês numérico exige separador para "48H" não virar 04/08.
_DATA = (
    r"(?P<d>\d{1,2})(?:\s*(?:DE\s+|/|-|\.)?\s*"
    r"(?P<ma>JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)[A-Z]*"
    r"|\s*[/.-]\s*(?P<mn>\d{1,2})(?!\d))"
    r"(?:\s*(?:DE\s+|/|-|\.)?\s*(?P<a>\d{4}|\d{2}))?(?!\d)"
)
_REL = r"(?:(?P<h>\d{1,3})\s*H(?:ORAS?)?\b|(?P<du>\d{1,2})\s*DIAS?\s+UTEIS)"

_GATILHOS = (
    ("empenho", r"(?:PRAZO\s+(?:PARA|DE)?\s*(?:O\s+)?EMPENHO|EMPENH[A-Z]*|EMPH|EMP\.?)"
                r"\s*(?:EM\s+)?(?:ATE\b|:)\s*:?\s*(?:ATE\s+)?"),
    ("liquidacao", r"LIQUIDA[A-Z]*\s*(?:ATE\b|:)\s*:?\s*"),
    ("recolhimento", r"RECOLHER\s+(?:OS\s+)?SALDOS?\s*(?:ATE\b|:)\s*:?\s*"),
)
_RE = {
    tipo: (re.compile(r"\b" + g + _REL), re.compile(r"\b" + g + _DATA))
    for tipo, g in _GATILHOS
}
_RE_IMEDIATO = re.compile(
    r"(?:PRAZO\s+(?:PARA|DE)?\s*(?:O\s+)?EMPENHO\s*:?\s*|EMPENH[A-Z]*\s+|EMPH\s+)(?:IMEDIATO|IMTO)\b"
)


def _normalizar(txt):
    t = unicodedata.normalize("NFKD", str(txt or "").upper())
    return "".join(c for c in t if not unicodedata.combining(c))


def _somar_dias_uteis(base, n):
    d = base
    while n > 0:
        d += datetime.timedelta(days=1)
        if d.weekday() < 5:
            n -= 1
    return d


def _data_br(s):
    try:
        return datetime.datetime.strptime(str(s).strip(), "%d/%m/%Y").date()
    except (TypeError, ValueError):
        return None


def extrair_prazo_nc(texto, dia_emissao=None):
    """Devolve o prazo descrito na NC ou None.

    texto: descrição da NC. dia_emissao: "dd/mm/aaaa" (base dos prazos relativos e do ano
    omitido). Prioridade: empenho > liquidação > recolhimento; dentro do tipo, o primeiro.

    Retorno: {"data": "dd/mm/aaaa" | None, "tipo": "empenho"|"liquidacao"|"recolhimento",
              "imediato": bool, "regra": str, "trecho": str}
    """
    t = _normalizar(texto)
    if not t:
        return None
    emissao = _data_br(dia_emissao)

    m = _RE_IMEDIATO.search(t)
    if m:
        return {"data": None, "tipo": "empenho", "imediato": True,
                "regra": "imediato (texto da NC)", "trecho": m.group(0).strip()}

    for tipo, (re_rel, re_data) in _RE.items():
        cands = []
        for m in re_data.finditer(t):
            g = m.groupdict()
            try:
                mes = int(g["mn"]) if g["mn"] else _MESES[g["ma"][:3]]
                if g["a"]:
                    ano = int(g["a"]) + (2000 if len(g["a"]) == 2 else 0)
                    regra = "data explícita na NC"
                else:
                    ano = emissao.year if emissao else datetime.date.today().year
                    regra = "data na NC (ano da emissão)"
                dt = datetime.date(ano, mes, int(g["d"]))
            except (KeyError, ValueError):
                continue
            cands.append((m.start(), dt, regra, m.group(0)))
        for m in re_rel.finditer(t):
            if not emissao:
                continue
            g = m.groupdict()
            if g["h"]:
                horas = int(g["h"])
                dt = emissao + datetime.timedelta(days=-(-horas // 24))
                regra = f"{horas}h após a emissão da NC ({emissao:%d/%m/%Y})"
            else:
                n = int(g["du"])
                dt = _somar_dias_uteis(emissao, n)
                regra = f"{n} dias úteis após a emissão da NC ({emissao:%d/%m/%Y}), sem feriados"
            cands.append((m.start(), dt, regra, m.group(0)))
        if cands:
            _, dt, regra, trecho = min(cands, key=lambda c: c[0])
            return {"data": dt.strftime("%d/%m/%Y"), "tipo": tipo, "imediato": False,
                    "regra": regra, "trecho": trecho.strip()}
    return None


_ROTULO_TIPO = {"empenho": "empenho", "liquidacao": "liquidação", "recolhimento": "recolhimento"}


def resolver_prazo(texto, dia_emissao=None, curado=None):
    """Prazo para exibição. Fonte: texto da NC (SIAFI) > resumos_nc.json > não informado.

    Devolve (prazo_str, info). prazo_str é "dd/mm/aaaa", "IMEDIATO" ou "Não informado na NC".
    info traz tipo, fonte, regra e, se houver, o prazo curado divergente.
    """
    p = extrair_prazo_nc(texto, dia_emissao)
    curado_norm = None
    if curado:
        c = str(curado).strip()
        dc = _data_br(c) or (_data_br(c[:6] + "20" + c[6:]) if re.match(r"^\d{2}/\d{2}/\d{2}$", c) else None)
        curado_norm = dc.strftime("%d/%m/%Y") if dc else c
    if p:
        prazo = "IMEDIATO" if p["imediato"] else p["data"]
        info = {"tipo": _ROTULO_TIPO[p["tipo"]], "fonte": "descrição da NC (SIAFI)",
                "regra": p["regra"], "trecho": p["trecho"]}
        if curado_norm and curado_norm != prazo:
            info["curado_divergente"] = curado_norm
        return prazo, info
    if curado_norm:
        return curado_norm, {"tipo": "empenho", "fonte": "resumos_nc.json (curadoria)",
                             "regra": "registro manual", "trecho": ""}
    return "Não informado na NC", {"tipo": "", "fonte": "", "regra": "texto da NC sem prazo", "trecho": ""}
