import win32com.client as win32
from PIL import ImageGrab
import os
import time
import psutil
from datetime import datetime

# ✅ Fecha o Excel completamente (força fechamento)
def matar_excel():
    for processo in psutil.process_iter(attrs=["pid", "name"]):
        if processo.info["name"] and "EXCEL.EXE" in processo.info["name"].upper():
            try:
                os.kill(processo.info["pid"], 9)  # Mata o processo do Excel
                print("⚠ Excel fechado com sucesso.")
                time.sleep(3)  # Aguarda o encerramento
            except Exception as e:
                print(f"❌ Erro ao fechar o Excel: {e}")

# ✅ Captura imagem de um intervalo no Excel
def capturar_imagem_excel(arquivo_excel, planilha_nome, celulas, nome_imagem):
    matar_excel()  # Fecha qualquer instância do Excel antes de iniciar
    excel = None
    workbook = None

    try:
        # 🔹 Abre o Excel
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False  # Não exibe o Excel
        time.sleep(3)  # Aguarda abertura do Excel

        workbook = excel.Workbooks.Open(arquivo_excel)
        time.sleep(3)  # Espera o arquivo carregar

        # 🔹 Verifica se a planilha existe
        planilhas = [sheet.Name for sheet in workbook.Worksheets]
        if planilha_nome not in planilhas:
            raise ValueError(f"❌ Planilha '{planilha_nome}' não encontrada! Disponíveis: {planilhas}")

        sheet = workbook.Worksheets(planilha_nome)
        sheet.Activate()
        time.sleep(2)  # Espera ativação

        # 🔹 Captura a imagem da célula especificada
        sheet.Range(celulas).CopyPicture(Format=win32.constants.xlBitmap)
        time.sleep(2)  # Aguarda cópia para a área de transferência

        # 🔹 Aguarda a imagem aparecer na área de transferência (máx. 10s)
        imagem = None
        for i in range(10):
            imagem = ImageGrab.grabclipboard()
            if imagem:
                break
            print(f"⏳ Aguardando captura da imagem... ({i+1}/10)")
            time.sleep(1)

        if not imagem:
            raise ValueError("❌ Falha ao capturar a imagem do Excel!")

        imagem.save(nome_imagem)
        print(f"✅ Imagem salva com sucesso: {nome_imagem}")

    except Exception as e:
        print(f"❌ Erro ao capturar imagem do Excel: {e}")

    finally:
        # 🔹 Fecha o Excel corretamente
        if workbook:
            try:
                workbook.Close(SaveChanges=False)
                time.sleep(2)
            except Exception as e:
                print(f"❌ Erro ao fechar a planilha: {e}")

        if excel:
            try:
                excel.Quit()
                time.sleep(2)
            except Exception as e:
                print(f"❌ Erro ao fechar o Excel: {e}")

        # 🔹 Garante que o Excel foi fechado
        matar_excel()

# ✅ Função para enviar e-mail com as imagens capturadas
def enviar_email():
    matar_excel()  # Fecha qualquer Excel antes de iniciar

    arquivo_excel = r".xlsx" #formato final do execel precisar se (xlsx) e informar caminho localizado do arquivo em xlsx
    nome_planilha = "Nome da aba da planilha" #Informar nome da aba que está na planilha do execel que irá ser printada as fotos

    # Caminhos para salvar as imagens
    imagens = {
        "imagem1": os.path.join(os.getcwd(), "imagem1.png"),
        "imagem2": os.path.join(os.getcwd(), "imagem2.png"),
        "imagem3": os.path.join(os.getcwd(), "imagem3.png"),
    }

    # Captura imagens da planilha
    capturar_imagem_excel(arquivo_excel, nome_planilha, "A2:C15", imagens["imagem1"]) # aqui está buscando capturar as imagens de cada tabela correspondente
    capturar_imagem_excel(arquivo_excel, nome_planilha, "E2:O17", imagens["imagem2"])
    capturar_imagem_excel(arquivo_excel, nome_planilha, "Q2:X6", imagens["imagem3"])

    data_atual = datetime.now().strftime("%d/%m/%Y")

    try:
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)

        # Definindo destinatários
        email.To = ""#Emeail para quem irá ser enviado
        email.CC = ""#Emeail de copia informando para quem irá ser enviado

        # Definindo o assunto e o corpo do e-mail
        email.Subject = f"📊 Status Report Testes [{data_atual}] - Regressão Ativação Simplificada"
        email.HTMLBody = f"""
        <html>
        <body>
            <p>Boa tarde!</p>
            <p>Segue status report  {data_atual}</p>
            <p><img src="cid:Imagem1"></p>
            <p><strong>Evolução detalhada das tasks:</strong></p>
            <p><img src="cid:Imagem2"></p>
            <p><strong>Bugs x Criticidade:</strong></p>
            <p><img src="cid:Imagem3"></p>
            <p><strong>Observação / Atividades Complementares:</strong></p>
        </body>
        </html>
        """

        # ✅ Anexa imagens apenas se existirem
        for nome, caminho in imagens.items():
            if os.path.exists(caminho):
                anexo = email.Attachments.Add(caminho)
                anexo.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", nome.capitalize())
                print(f"📎 {nome} anexado com sucesso!")

        email.Display()  # Exibe o e-mail antes do envio
        print("📧 E-mail pronto para envio!")

    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")

# ✅ Executa o envio do e-mail
if __name__ == "__main__":
    enviar_email()
