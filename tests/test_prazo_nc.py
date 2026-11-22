import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from prazo_nc import extrair_prazo_nc, resolver_prazo  # noqa: E402

CASO_TELA = ("25PDREXTR032* SOLICITACAO: 202201000369 - REPARACAO / INSTALACAO ELETRICA / PAVILHAO "
             "OFICINA / BCMS. EMPH CFM MSG SIAFI 2026/3272070-DEC, DE 26 FEV 26. NC 4976. PROIBIDO "
             "DETALHAMENTO DE CREDITO. PRAZO PARA EMPENHO: 06/11/2026.")


class TestExtrairPrazo(unittest.TestCase):
    def data(self, txt, emissao="10/03/2026"):
        p = extrair_prazo_nc(txt, emissao)
        return p and (p["data"], p["tipo"])

    def test_caso_relatado_nc_dec(self):
        # A NC do print: o texto diz 06/11/2026; a tela mostrava 15/12/2026 (inventado).
        self.assertEqual(self.data(CASO_TELA, "05/11/2026"), ("06/11/2026", "empenho"))
        # "DE 26 FEV 26" é data do documento, não prazo
        self.assertNotEqual(self.data(CASO_TELA)[0], "26/02/2026")

    def test_formatos_de_data(self):
        casos = {
            "EMPENHAR ATÉ 30/MAR/26.": "30/03/2026",
            "EMPENHO ATÉ 30SET2026": "30/09/2026",
            "EMPENHAR ATE 15OUT26": "15/10/2026",
            "EMPENHO ATE 30 OUT 26": "30/10/2026",
            "EMPH ATE 15 MAIO 26": "15/05/2026",
            "EMPH ATE: 15/10/26": "15/10/2026",
            "EMPH ATE 20/11/2026": "20/11/2026",
            "EMPENHO ATÉ 30NOV2026. LIQUIDAR ATE 30DEZ26": "30/11/2026",
        }
        for txt, esperado in casos.items():
            with self.subTest(txt=txt):
                self.assertEqual(self.data(txt)[0], esperado)

    def test_ano_omitido_usa_ano_da_emissao(self):
        self.assertEqual(self.data("EMPENHAR ATE 30SET", "02/02/2025")[0], "30/09/2025")

    def test_prazos_relativos(self):
        # 48h a partir da emissão (terça 10/11/2026)
        self.assertEqual(self.data("EMPH ATE 48H", "10/11/2026")[0], "12/11/2026")
        # 7 dias úteis a partir de sexta 13/11/2026 -> 24/11/2026
        self.assertEqual(self.data("EMPH ATE 7 DIAS UTEIS", "13/11/2026")[0], "24/11/2026")
        self.assertEqual(self.data("EMPH EM ATE 5 DIAS UTEIS - EMP C", "27/04/2026")[0], "04/05/2026")
        # sem data de emissão, prazo relativo não vira data
        self.assertIsNone(extrair_prazo_nc("EMPH ATE 48H", None))

    def test_48h_nao_vira_data(self):
        self.assertNotEqual(self.data("EMPH ATE 48H", "10/11/2026")[0], "04/08/2026")

    def test_imediato(self):
        for txt in ("PRAZO DE EMPENHO: IMEDIATO", "EMPENHO IMEDIATO", "EMPH IMEDIATO", "PRAZO EMPENHO: IMEDIATO"):
            with self.subTest(txt=txt):
                p = extrair_prazo_nc(txt, "10/11/2026")
                self.assertTrue(p["imediato"])
                self.assertIsNone(p["data"])

    def test_liquidacao_e_recolhimento(self):
        self.assertEqual(self.data("LIQUIDAR ATÉ 30 SET 26"), ("30/09/2026", "liquidacao"))
        self.assertEqual(self.data("RECOLHER SALDOS ATE 01 NOV 26 - DESCE"), ("01/11/2026", "recolhimento"))
        # empenho tem prioridade mesmo aparecendo depois
        self.assertEqual(self.data("LIQUIDAR ATE 30DEZ26. EMPENHAR ATE 30NOV26"), ("30/11/2026", "empenho"))

    def test_sem_prazo_nao_inventa(self):
        for txt in ("", None, "AQUISICAO DE MATERIAL ELETRICO. REF DIEX 123 DE 05/05/2026",
                    "PRAZO DE APLICACAO CONFORME PDR", "SISTEMA TEMPORARIO ATE 30NOV26"):
            with self.subTest(txt=txt):
                self.assertIsNone(extrair_prazo_nc(txt, "10/11/2026"))


class TestResolverPrazo(unittest.TestCase):
    def test_texto_da_nc_vence_curadoria(self):
        prazo, info = resolver_prazo(CASO_TELA, "05/11/2026", curado="15/12/26")
        self.assertEqual(prazo, "06/11/2026")
        self.assertEqual(info["fonte"], "descrição da NC (SIAFI)")
        self.assertEqual(info["curado_divergente"], "15/12/2026")

    def test_curadoria_quando_texto_nao_tem(self):
        prazo, info = resolver_prazo("AQUISICAO DE INSUMOS", "05/01/2026", curado="31/03/26")
        self.assertEqual(prazo, "31/03/2026")
        self.assertIn("curadoria", info["fonte"])

    def test_sem_fonte(self):
        prazo, _ = resolver_prazo("AQUISICAO DE INSUMOS", "05/01/2026")
        self.assertEqual(prazo, "Não informado na NC")


if __name__ == "__main__":
    unittest.main()
