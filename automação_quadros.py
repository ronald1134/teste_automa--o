import os
import shutil
import time
import datetime
import win32com.client

# Novos caminhos das pastas
pasta_quadros = r"G:\Drives compartilhados\Alliar BF - Planejamento\QUADROS\4 - Quadro de Colaboradores - TESTE\2025\02. Fevereiro"
pasta_update = r"G:\Drives compartilhados\Alliar BF - Planejamento\00 - BASES DNA\02.BASES\ARQUIVOS_BASE\UPDATE - TESTE"
pasta_bases = r"G:\Drives compartilhados\Contact Center - Planejamento\01 - DashBoards\02 - IGP\00 - BASES - TESTE\2025\2025.02"

# Mapeamento dos arquivos para substituição
arquivos_quadros_mapping = {
    "AXIAL_FEVEREIRO25.xlsx": "AXIAL_QUADRO.xlsx",
    "CEDIM_FEVEREIRO25.xlsx": "CEDIM_QUADRO.xlsx",
    "DELFIN_FEVEREIRO25.xlsx": "DELFIN_QUADRO.xlsx",
    "GCO_FEVEREIRO25.xlsx": "GCO_QUADRO.xlsx",
    "MULTI_FEVEREIRO25.xlsx": "MSCAN_QUADRO.xlsx",
    "PLANI_FEVEREIRO25.xlsx": "PLANI_PROIM_QUADRO.xlsx",
    "SABED_FEVEREIRO25.xlsx": "SABEDOTTI_QUADRO.xlsx",
    "SP_FEVEREIRO_25.xlsx": "QUADRO.xlsx",  # Arquivo especial
    "SOM_FEVEREIRO25.xlsx": "SOM_QUADRO.xlsx",
    "UMDI_FEVEREIRO25.xlsx": "UMDI_QUADRO.xlsx"
}

# Inicializar Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # Rodar em segundo plano
excel.DisplayAlerts = False  # Desativar alertas

# Data atual para comparação
hoje = datetime.datetime.now().date()

# Passo 1: Abrir cada arquivo na pasta "quadros de colaboradores" e salvar na "UPDATE" (substituindo os existentes)
for arquivo_origem, arquivo_destino in arquivos_quadros_mapping.items():
    origem = os.path.join(pasta_quadros, arquivo_origem)
    destino = os.path.join(pasta_update, arquivo_destino)

    if os.path.exists(origem):
        try:
            wb = excel.Workbooks.Open(origem)
            wb.SaveAs(destino)
            wb.Close(SaveChanges=True)
            print(f"Salvo e substituído na UPDATE: {arquivo_destino}")
        except Exception as e:
            print(f"Erro ao salvar {arquivo_origem}: {e}")
    else:
        print(f"Arquivo não encontrado: {arquivo_origem}")

# Passo 2: Copiar arquivos da "UPDATE" para "00-BASES/2025/2025.02"
arquivos_update = [f for f in os.listdir(pasta_update) if f.endswith(".xlsx")]

for arquivo in arquivos_update:
    origem = os.path.join(pasta_update, arquivo)
    destino = os.path.join(pasta_bases, arquivo)
    shutil.copy2(origem, destino)  # copy2 mantém a data original
    print(f"Copiado para 00-BASES: {arquivo}")

# Passo 3: Atualizar arquivos da 00-BASES apenas se estiverem desatualizados
arquivos_bases = [f for f in os.listdir(pasta_bases) if f.endswith(".xlsx")]
arquivos_bases.sort(key=lambda x: os.path.getmtime(os.path.join(pasta_bases, x)))  # Ordena por modificação

for arquivo in arquivos_bases:
    arquivo_path = os.path.join(pasta_bases, arquivo)
    data_modificacao = datetime.datetime.fromtimestamp(os.path.getmtime(arquivo_path)).date()

    # Se o arquivo foi modificado antes de hoje, atualizar
    if data_modificacao < hoje:
        try:
            wb = excel.Workbooks.Open(arquivo_path)
            wb.RefreshAll()  # Atualiza todas as conexões e consultas
            time.sleep(5)  # Aguarda atualização
            wb.Save()
            wb.Close(SaveChanges=True)
            print(f"Atualizado: {arquivo}")
        except Exception as e:
            print(f"Erro ao atualizar {arquivo}: {e}")
    else:
        print(f"Arquivo já atualizado: {arquivo}")

# Fechar Excel
excel.Quit()
print("Processo concluído.")
