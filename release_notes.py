import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Configurações para autenticação
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("rn-automation-424115-8c73ed877cf8.json", scope)
client = gspread.authorize(creds)

# Solicita ao usuário a URL da planilha e o nome da aba desejada
spreadsheet_url = input("Por favor, insira a URL da planilha: ")
aba_desejada = input("Por favor, insira o nome da aba desejada: ")

# Obtém a planilha diretamente pelo URL
spreadsheet = client.open_by_url(spreadsheet_url)

# Procura pela aba desejada na lista de abas
worksheet = None
for ws in spreadsheet.worksheets():
    if ws.title == aba_desejada:
        worksheet = ws
        break

if worksheet is None:
    raise ValueError(f"A aba '{aba_desejada}' não foi encontrada na planilha")

# Obtém os índices das colunas desejadas
header_row = worksheet.row_values(1)
column_names = [
    "Model",
    "SPL\n(Security Patch Level)",
    "IMEI SV \n(SVN) Number",
    "Software TA",
    "Fingerprint",
    "Modem Release",
    "FSG Release",
    "Bootloader"
]

column_indices = {}
for name in column_names:
    try:
        column_indices[name] = header_row.index(name) + 1  # gspread usa índices começando em 1
    except ValueError:
        raise ValueError(f"A coluna '{name}' não foi encontrada na aba '{aba_desejada}'")

# Obtém todos os valores das colunas desejadas
column_values = {name: worksheet.col_values(idx) for name, idx in column_indices.items()}

# Verifica se todas as colunas têm o mesmo número de linhas e ajusta se necessário
max_length = max(len(values) for values in column_values.values())
for name in column_values:
    column_values[name] += [""] * (max_length - len(column_values[name]))

# Constrói uma lista de linhas a partir dos valores das colunas
rows = []
for i in range(max_length):
    row = [column_values[name][i] for name in column_names]
    rows.append(row)

# Remove linhas duplicadas
unique_rows = []
seen_rows = set()
for row in rows:
    row_tuple = tuple(row)
    if row_tuple not in seen_rows:
        seen_rows.add(row_tuple)
        unique_rows.append(row)

# Salva os valores únicos em um arquivo .txt em formato de tabela
with open("output.txt", "w") as file:
    headers = "\t".join(column_names)
    file.write(headers + "\n")
    for row in unique_rows:
        row_str = "\t".join(row)
        file.write(row_str + "\n")

print("Valores das colunas especificadas da aba especificada salvos em output.txt, com linhas duplicadas removidas")
