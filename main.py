# -*- coding: utf-8 -*-
from __future__ import print_function
import os
import os.path
# from reportlab.pdfgen import canvas
# from reportlab.lib.pagesizes import A4pi
# from pygments.styles import get_all_styles
import pandas as pd
import shutil
import subprocess
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError



SCOPES = ['https://www.googleapis.com/auth/spreadsheets']



def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None






    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES) 

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId='1Mk0uLFHCwrybovkp2aPt_IS7EkQK3gMgtd9',
                                    range='relestendido!A1:Z3000').execute()

        values = result.get('values', [])

        if not values:
            print('No data found.')
            return

        else:
                df = pd.DataFrame(values)
                df.columns = df.iloc[0]


        genero = df.genero
        especie = df.especie
        # figura = df.figuras[1]


        codigo_do_relatorio= df.cod_rel[1]
        endereco = df.endereco[1]
        nomecliente = df.nomecliente[1]
        os_or = df.os_or[1]

        texto_objetivos = []
        texto_objetivos.append(df.objetivos_do_projeto[1])

        texto_resultados = []
        texto_resultados.append(df.resultados[1])

        texto_discussoes = []
        texto_discussoes.append(df.discussoes[1])

        texto_sugestoes_de_acoes = []
        texto_sugestoes_de_acoes.append(df.sugestoes_de_acoes[1])

        texto_conclusao = []
        texto_conclusao.append(df.conclusoes[1])


        texto_objetivos = ",".join(texto_objetivos)
        texto_objetivos = texto_objetivos.split()
        texto_final_objetivos = texto_objetivos  #igualar em todos


        for termo in texto_objetivos:
            for gender in genero:
                if termo == gender:
                    a = texto_objetivos.index(termo)
                    texto_objetivos[a] = "\\textit{" + termo + "}"

        for substituto in texto_objetivos:
            for specie in especie:
                if substituto == specie:
                    b = texto_objetivos.index(substituto)
                    texto_objetivos[b] = "\\textit{" + substituto + "}"
                    texto_final_objetivos = " ".join(texto_objetivos)
                    texto_final_objetivos = texto_final_objetivos.replace(",", " ,")
                    texto_final_objetivos = texto_final_objetivos.replace(".", " .")
                    texto_final_objetivos = texto_final_objetivos.replace(";", " ;")
                    texto_final_objetivos = texto_final_objetivos.split()

        for newspecie in texto_final_objetivos:
            for newItembanking in especie:
                if newspecie == newItembanking:
                    c = texto_final_objetivos.index(newspecie)
                    texto_final_objetivos[c] = "\\textit{" + newspecie + "}"

        texto_final_objetivos = " ".join(texto_final_objetivos)
        texto_final_objetivos = texto_final_objetivos.replace(" ,", ",")
        texto_final_objetivos = texto_final_objetivos.replace(" .", ".")
        texto_final_objetivos = texto_final_objetivos.replace(" ;", ";")


        texto_resultados = ",".join(texto_resultados)
        texto_resultados = texto_resultados.split()
        texto_final_resultados = texto_resultados  # igualar em todos

        for argumento in texto_resultados:
            for substantivo in genero:
                if argumento == substantivo:
                    r = texto_resultados.index(argumento)
                    texto_resultados[r] = "\\textit{" + argumento + "}"

        for substitute in texto_resultados:
            for subjetivo in especie:
                if substitute == subjetivo:
                    t = texto_resultados.index(substitute)
                    texto_resultados[t] = "\\textit{" + substitute + "}"
                    texto_final_resultados = " ".join(texto_resultados)
                    texto_final_resultados = texto_final_resultados.replace(",", " ,")
                    texto_final_resultados = texto_final_resultados.replace(".", " .")
                    texto_final_resultados = texto_final_resultados.replace(";", " ;")
                    texto_final_resultados = texto_final_resultados.split()

        for newspecifico in texto_final_resultados:
            for newItembankingnew in especie:
                if newspecifico == newItembankingnew:
                    s = texto_final_resultados.index(newspecifico)
                    texto_final_resultados[s] = "\\textit{" + newspecifico + "}"

        texto_final_resultados = " ".join(texto_final_resultados)
        texto_final_resultados = texto_final_resultados.replace(" ,", ",")
        texto_final_resultados = texto_final_resultados.replace(" .", ".")
        texto_final_resultados = texto_final_resultados.replace(" ;", ";")



        texto_discussoes = ",".join(texto_discussoes)
        texto_discussoes = texto_discussoes.split()
        texto_final_discussoes= texto_discussoes

        for drow in texto_discussoes:
            for generico in genero:
                if drow == generico:
                    g = texto_discussoes.index(drow)
                    texto_discussoes[g] = "\\textit{" + drow + "}"

        for substr in texto_discussoes:
            for specimen in especie:
                if substr == specimen:
                    h = texto_discussoes.index(substr)
                    texto_discussoes[h] = "\\textit{" + substr + "}"
                    texto_final_discussoes = " ".join(texto_discussoes)
                    texto_final_discussoes = texto_final_discussoes.replace(",", " ,")
                    texto_final_discussoes = texto_final_discussoes.replace(".", " .")
                    texto_final_discussoes = texto_final_discussoes.replace(";", " ;")
                    texto_final_discussoes = texto_final_discussoes.split()


        for newspecimen in texto_final_discussoes:
            for novoItem in especie:
                if novoItem == newspecimen:
                    j = texto_final_discussoes.index(newspecimen)
                    texto_final_discussoes[j] = "\\textit{" + newspecimen + "}"

        texto_final_discussoes = " ".join(texto_final_discussoes)
        texto_final_discussoes = texto_final_discussoes.replace(" ,", ",")
        texto_final_discussoes = texto_final_discussoes.replace(" .", ".")
        texto_final_discussoes = texto_final_discussoes.replace(" ;", ";")


        texto_sugestoes_de_acoes = ",".join(texto_sugestoes_de_acoes)
        texto_sugestoes_de_acoes = texto_sugestoes_de_acoes.split()
        texto_final_sugestoes_de_acoes = texto_sugestoes_de_acoes

        for argumentoz in texto_sugestoes_de_acoes:
            for generation in genero:
                if argumentoz == generation:
                    o = texto_sugestoes_de_acoes.index(argumentoz)
                    texto_sugestoes_de_acoes[o] = "\\textit{" + argumentoz + "}"

        for substrato in texto_sugestoes_de_acoes:
            for special in especie:
                if substrato == special:
                    p = texto_sugestoes_de_acoes.index(substrato)
                    texto_sugestoes_de_acoes[p] = "\\textit{" + substrato + "}"
                    texto_final_sugestoes_de_acoes = " ".join(texto_sugestoes_de_acoes)
                    texto_final_sugestoes_de_acoes = texto_final_sugestoes_de_acoes.replace(",", " ,")
                    texto_final_sugestoes_de_acoes = texto_final_sugestoes_de_acoes.replace(".", " .")
                    texto_final_sugestoes_de_acoes = texto_final_sugestoes_de_acoes.replace(";", " ;")
                    texto_final_sugestoes_de_acoes = texto_final_sugestoes_de_acoes.split()

        for newspecimenz in texto_final_sugestoes_de_acoes:
            for novoItemz in especie:
                if novoItemz == newspecimenz:
                    q = texto_final_sugestoes_de_acoes.index(newspecimenz)
                    texto_final_sugestoes_de_acoes[q] = "\\textit{" + newspecimenz + "}"

        texto_final_sugestoes_de_acoes = " ".join(texto_final_sugestoes_de_acoes)
        texto_final_sugestoes_de_acoes = texto_final_sugestoes_de_acoes.replace(" ,", ",")
        texto_final_sugestoes_de_acoes = texto_final_sugestoes_de_acoes.replace(" .", ".")
        texto_final_sugestoes_de_acoes = texto_final_sugestoes_de_acoes.replace(" ;", ";")



        texto_conclusao = ",".join(texto_conclusao)
        texto_conclusao = texto_conclusao.split()
        texto_final = texto_conclusao


        for palavra in texto_conclusao:
            for gen in genero:
                if palavra == gen:
                    x = texto_conclusao.index(palavra)
                    texto_conclusao[x] = "\\textit{" + palavra + "}"

        for subs in texto_conclusao:
            for spp in especie:
                if subs == spp:
                    y = texto_conclusao.index(subs)
                    texto_conclusao[y] = "\\textit{" + subs + "}"
                    texto_final = " ".join(texto_conclusao)
                    texto_final = texto_final.replace(",", " ,")
                    texto_final = texto_final.replace(".", " .")
                    texto_final = texto_final.replace(";", " ;")
                    texto_final = texto_final.split()

        for newspp in texto_final:
            for newItembank in especie:
                if newspp == newItembank:
                    z = texto_final.index(newspp)
                    texto_final[z] = "\\textit{" + newspp + "}"

        texto_final = " ".join(texto_final)
        texto_final = texto_final.replace(" ,", ",")
        texto_final = texto_final.replace(" .", ".")
        texto_final = texto_final.replace(" ;", ";")




        diretorio_original = Path("./relestendido")
        diretorio_temporario = Path("./relestentidocopia")

        if diretorio_temporario.is_dir():
            shutil.rmtree(diretorio_temporario)
        shutil.copytree(diretorio_original, diretorio_temporario)


        capa = diretorio_temporario/ "main.tex"
        coisas = capa.read_text()
        coisas = coisas.replace("$CODIGO-RELATORIO", codigo_do_relatorio)
        capa.write_text(coisas)

        capa = diretorio_temporario/ "capa.tex"
        coisass = capa.read_text()
        coisass = coisass.replace("$NOME-CLIENTE", nomecliente )
        capa.write_text(coisass)

        capa = diretorio_temporario/ "capa.tex"
        coisasss = capa.read_text()
        coisasss = coisasss.replace("$ENDERECO", endereco)
        capa.write_text(coisasss)

        capa = diretorio_temporario / "capa.tex"
        coisassss = capa.read_text()
        coisassss = coisassss.replace("$os_or", os_or)
        capa.write_text(coisassss)

        objetivos = diretorio_temporario / "objetivos_do_projeto.tex"
        conjunto = objetivos.read_text(encoding="UTF8")
        conjunto = conjunto.replace("$Objetivos", texto_final_objetivos)
        objetivos.write_text(conjunto, encoding="UTF8")

        resultados = diretorio_temporario / "resultados.tex"
        trecho = resultados.read_text(encoding="UTF8")
        trecho = trecho.replace("$Resultados", texto_final_resultados)
        resultados.write_text(trecho, encoding="UTF8")

        discussoes = diretorio_temporario / "discussoes.tex"
        paragrafos = discussoes.read_text(encoding="UTF8")
        paragrafos = paragrafos.replace("$Discussoes", texto_final_discussoes)
        discussoes.write_text(paragrafos, encoding="UTF8")

        sugestoes_de_acoes = diretorio_temporario / "sugestoes_acoes.tex"
        escrito = sugestoes_de_acoes.read_text(encoding="UTF8")
        escrito = escrito.replace("$Sugestoes_de_acoes", texto_final_sugestoes_de_acoes)
        sugestoes_de_acoes.write_text(escrito, encoding="UTF8")

        conclusao = diretorio_temporario / "conclusoes.tex"
        text = conclusao.read_text(encoding="UTF8")
        text = text.replace("$Conclusao", texto_final)
        conclusao.write_text(text, encoding="UTF8")



        os.chdir(diretorio_temporario)
        cmd = ["pdflatex", "-interaction", "nonstopmode", "main.tex"]
        proc = subprocess.Popen(cmd)
        proc.communicate()

        if proc.returncode != 0:
            print("Erro ao gerar PDF!")

        # print('Name, Major:')
        # for row in values:
        #     # Print columns A and E, which correspond to indices 0 and 4.
        #     print('%s, %s' % (row[0], row[4]))
    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()


