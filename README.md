# Automação de Envio de Status Report no Microsoft Outlook

Este projeto é um script em Python desenvolvido para automatizar o processo de captura de imagens de intervalos específicos em uma planilha do Excel e enviá-las por e-mail como parte de um status report, utilizando o Microsoft Outlook. É ideal para a geração automática de relatórios de progresso e comunicação eficiente com a equipe.

## Funcionalidades

- Fecha quaisquer instâncias do Excel antes de iniciar o processo.
- Abre uma planilha do Excel e captura imagens de intervalos definidos.
- Salva as imagens capturadas no diretório atual.
- Envia um e-mail com as imagens anexadas para destinatários específicos usando o Outlook.

## Requisitos

- Python 3.x
- Bibliotecas Python:
  - `win32com.client` (parte do pacote `pywin32`)
  - `PIL` (Pillow)
  - `psutil`
- Microsoft Excel
- Microsoft Outlook

## Instalação

1. Clone o repositório para sua máquina local:

   ```bash
   git clone https://github.com/Ednavan/Status_Report_QA
   cd nome-do-repositorio
   ```text
2. Instale as bibliotecas necessárias:

   ```bash
   pip install pywin32 pillow psutil

Certifique-se de que o Excel e o Outlook estão instalados e configurados corretamente em seu sistema.


Uso
 - Atualize o caminho do arquivo Excel e o nome da planilha no script StatusReport.py.


 - Atualize os endereços de e-mail no script para os destinatários desejados.


Execute o script:

 - python StatusReport.py

 - O e-mail será preparado com as imagens anexadas e estará pronto para envio.



## Configuração do Ambiente
 - Excel: Certifique-se de que o arquivo Excel está disponível no caminho especificado e que a planilha contém os intervalos que você deseja capturar.
 - Outlook: Configure o Outlook com a conta de e-mail que você deseja usar para enviar os relatórios.

## Contribuições

- Faça um fork do projeto.

- Crie sua feature branch (git checkout -b feature/nova-feature).

- Commit suas mudanças (git commit -m 'Adiciona nova feature').

- Dê um push para a branch (git push origin feature/nova-feature).

- Abra um Pull Request.
