# %%
# Importar Bibliotecas
import pandas as pd
import os

# %%
# Pasta onde estão os arquivos
pasta = "data/arquivos"

# Lista de arquivos XLS/XLSX na pasta
arquivos = [f for f in os.listdir(pasta) if f.endswith((".xls", ".xlsx"))]

# Dicionário Placa -> Motorista
placa_motorista = {
    "XXX1111": "Motorista 1",
    "XXX2222": "Motorista 2",
    "XXX3333": "Motorista 3",
    "XXX4444": "Motorista 4",
    "XXX5555": "Motorista 5",
    "XXX6666": "Motorista 6",
    "XXX7777": "Motorista 7",
    "XXX8888": "Motorista 8",
    "XXX9999": "Motorista 9",
    "XXX1010": "Motorista 10",
    "XXX0011": "Motorista 11",
    "XXX1212": "Motorista 12",
    "XXX1313": "Motorista 13",
    "XXX1414": "Motorista 14",
    "XXX1515": "Motorista 15",
    "XXX1616": "Motorista 16"
}

# Colunas importantes
colunas_para_manter = [
    'Data/Hora início',
    'Placa',
    'Nr. de Frota',
    'Volume (l)',
    'Odometro',
    'Distância (km)',
    'Média (km/l)',
    'Frentista',
    'Motorista',
    'Cidade'
]

# Lista final com todos dados limpos
dados_todos = []

# %%
# LOOP EM TODOS ARQUIVOS
for arquivo in arquivos:
    caminho = os.path.join(pasta, arquivo)
    print(f"Processando: {arquivo}")

    try:
        # Ler planilha
        df = pd.read_excel(caminho)

        # Preencher motorista baseado na placa
        df["Motorista"] = df["Placa"].map(placa_motorista)

        # Filtrar apenas placas reconhecidas
        df = df.dropna(subset=["Motorista"])

        # Manter só colunas desejadas
        df = df[colunas_para_manter]

        # Adiciona na lista geral
        dados_todos.append(df)

    except Exception as e:
        print(f"❌ Erro ao processar {arquivo}: {e}")

# %%
# Junta tudo
if dados_todos:
    dados_novos = pd.concat(dados_todos, ignore_index=True)
else:
    raise Exception("Nenhum arquivo válido encontrado!")

# Caminho final
arquivo_excel = "data/Medias.xlsx"

# Se já existe, concatena com o existente
if os.path.exists(arquivo_excel):
    dados_existentes = pd.read_excel(arquivo_excel)
    dados_final = pd.concat([dados_existentes, dados_novos], ignore_index=True)
else:
    dados_final = dados_novos

# Remove duplicatas opcionais
# dados_final = dados_final.drop_duplicates()

# Salvar
dados_final.to_excel(arquivo_excel, index=False)

print("✔️ Processamento concluído! Arquivo salvo em Medias.xlsx")
