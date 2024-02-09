import os.path
import math
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def calcular_situacao(media, faltas):
    if faltas > 0.25 * 60: 
        return "Reprovado por Falta"
    elif media < 50:
        return "Reprovado por Nota"
    elif 50 <= media < 70:
        return "Exame Final"
    else:
        return "Aprovado"

def calcular_naf(media):
    naf = max(0, (50 * 2) - media)  
    return math.ceil(naf) 

def main():
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    else:
        print("Arquivo 'token.json' não encontrado.")
        return

    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()
        spreadsheet_id = "1nlo1snNRkceOu3bNBBSOnV-DS4Rmu9rpG5Pfs52uGvU"
        range_name = "engenharia_de_software!A4:H27"  

        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        values = result.get("values", [])

        if not values:
            print("Nenhum dado encontrado.")
            return

        updated_values = []
        for row in values:
            matricula = row[0]
            aluno = row[1]
            faltas = int(row[2]) if row[2] else 0
            p1 = float(row[3]) if row[3] else 0
            p2 = float(row[4]) if row[4] else 0
            p3 = float(row[5]) if row[5] else 0

            media = (p1 + p2 + p3) / 3

            situacao = calcular_situacao(media, faltas)

            naf = 0
            if situacao == "Exame Final":
                naf = calcular_naf(media)

            row[6] = situacao
            row[7] = naf

            updated_values.append(row)

        body = {"values": updated_values}
        result = sheet.values().update(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption="RAW",
            body=body,
        ).execute()

        print("valores atualizados")

    except HttpError as err:
        print(err)

if __name__ == "__main__":
    main()
