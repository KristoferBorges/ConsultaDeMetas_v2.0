import pandas as pd
import openpyxl
import numpy as np
from modulo import dateVerification
import subprocess


def abrir_arquivo(caminho_arquivo):
    try:
        # Abre o arquivo usando o programa padrão associado a ele no sistema operacional
        subprocess.Popen([caminho_arquivo], shell=True)
    except Exception as e:
        print(f"Ocorreu um erro ao abrir o arquivo: {e}")


caminho = "storage\listaRDMarcas.xlsx"
abrir_arquivo(caminho)
