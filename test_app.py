import os
import unittest
import pandas as pd
from unittest.mock import MagicMock, patch

import importlib.util

# Load module without instantiating Tk
spec = importlib.util.spec_from_file_location(
    "app_module", "Acessos-de-pessoas-na-academia-que-n-o-assinaram-o-termo.py"
)
app_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app_module)
AppCruzadorExcel = app_module.AppCruzadorExcel


class TestAppCruzadorExcel(unittest.TestCase):

    def setUp(self):
        # Create temporary test files
        self.test_dir = "test_files"
        os.makedirs(self.test_dir, exist_ok=True)

        self.csv_guia1_path = os.path.join(self.test_dir, "guia1.csv")
        self.csv_guia2_path = os.path.join(self.test_dir, "guia2.csv")
        self.excel_guia3_path = os.path.join(self.test_dir, "guia3.xlsx")

        # Guia 1 CSV data
        df1_data = pd.DataFrame([
            ["ID do Evento", "Horário", "Nome do Dispositivo", "Ponto do Evento", "ID Pessoal", "Nome", "Sobrenome", "Número do Cartão"],
            ["101", "2023-10-01 08:00", "Catraca 1", "Entrada", "P001", "", "", ""],
            ["102", "2023-10-01 09:00", "Catraca 2 -- s", "Saída", "P002", "", "", ""],
            ["103", "2023-10-01 10:00", "Catraca 1", "Entrada", "P003", "", "", ""]
        ])
        df1_data.to_csv(self.csv_guia1_path, index=False, header=False, sep="\t")

        # Guia 2 CSV data
        df2_data = pd.DataFrame([
            ["ID Pessoal", "Nome", "Sobrenome", "Número do Cartão"],
            ["P001", "João", "Silva", "CARD111"],
            ["P002", "Maria", "Santos", "CARD222"],
            ["P003", "Pedro", "Oliveira", "CARD333"]
        ])
        df2_data.to_csv(self.csv_guia2_path, index=False, header=False, sep="\t")

        # Guia 3 Excel data (contains CARD333 to exclude)
        df3_data = pd.DataFrame([
            ["Número do Cartão"],
            ["CARD333"]
        ])
        df3_data.to_excel(self.excel_guia3_path, index=False, header=False)

        # Mock App instance without Tk display requirement
        with patch.object(AppCruzadorExcel, "__init__", lambda self: None):
            self.app = AppCruzadorExcel()
            self.app.ACCENT_GREEN = "#27ae60"
            self.app.ACCENT_BLUE = "#4a6fa5"
            self.app.ACCENT_RED = "#e74c3c"
            self.app.TEXT_MUTED = "#a0a0b0"
            self.app.tab4 = MagicMock()
            self.app.df_guia1 = None
            self.app.df_guia2 = None
            self.app.df_guia3 = None
            self.app.df_resultado = None

    def tearDown(self):
        # Remove test files
        for f in [self.csv_guia1_path, self.csv_guia2_path, self.excel_guia3_path]:
            if os.path.exists(f):
                os.remove(f)
        if os.path.exists(self.test_dir):
            os.rmdir(self.test_dir)

    def test_ler_arquivo_df_csv(self):
        df_raw = self.app._ler_arquivo_df(self.csv_guia1_path)
        self.assertIsNotNone(df_raw)
        self.assertFalse(df_raw.empty)
        self.assertEqual(df_raw.iloc[0, 0], "ID do Evento")

    def test_ler_arquivo_df_excel(self):
        df_raw = self.app._ler_arquivo_df(self.excel_guia3_path)
        self.assertIsNotNone(df_raw)
        self.assertFalse(df_raw.empty)
        self.assertEqual(df_raw.iloc[0, 0], "Número do Cartão")

    def test_tratar_cabecalhos(self):
        df_raw = self.app._ler_arquivo_df(self.csv_guia1_path)
        df_tratado = self.app._tratar_cabecalhos(df_raw, 1)
        self.assertIn("ID do Evento", df_tratado.columns)
        self.assertIn("ID Pessoal", df_tratado.columns)

    def test_processar_dados(self):
        # Setup dataframes
        df1_raw = self.app._ler_arquivo_df(self.csv_guia1_path)
        self.app.df_guia1 = self.app._tratar_cabecalhos(df1_raw, 1)

        df2_raw = self.app._ler_arquivo_df(self.csv_guia2_path)
        self.app.df_guia2 = self.app._tratar_cabecalhos(df2_raw, 2)

        df3_raw = self.app._ler_arquivo_df(self.excel_guia3_path)
        self.app.df_guia3 = self.app._tratar_cabecalhos(df3_raw, 3)

        # Mock Treeview, buttons and labels to prevent TK GUI errors during _processar_dados
        self.app.tree_res = MagicMock()
        self.app.lbl_res_summary = MagicMock()
        self.app.btn_exportar = MagicMock()
        self.app.notebook = MagicMock()

        with patch("tkinter.messagebox.showinfo"), patch("tkinter.messagebox.showerror"):
            self.app._processar_dados()

        self.assertIsNotNone(self.app.df_resultado)
        # P001 -> CARD111 (kept)
        # P002 -> CARD222 (kept)
        # P003 -> CARD333 (filtered out / excluded by Guia 3)
        # Expected final count: 2
        self.assertEqual(len(self.app.df_resultado), 2)
        cartoes_finais = list(self.app.df_resultado["Número do Cartão"])
        self.assertIn("CARD111", cartoes_finais)
        self.assertIn("CARD222", cartoes_finais)
        self.assertNotIn("CARD333", cartoes_finais)


if __name__ == "__main__":
    unittest.main()
