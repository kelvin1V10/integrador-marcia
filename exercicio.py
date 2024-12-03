import pandas as pd
import smtplib
from email.mime.text import MIMEText
from datetime import datetime
 
 
def disparar_email_aviso(mensagem_aviso):
    remetente_email = "mtabrasil904@gmail.com"  # Conta de e-mail de origem
    destinatario_email = "mtabrasil904@gmail.com"  # Conta de e-mail de destino
    senha_acesso = "pxbq joop txcq vbxb"  # App Password ou senha do e-mail
 
    # Configuração do conteúdo do e-mail
    conteudo_email = MIMEText(mensagem_aviso, _charset="utf-8")
    conteudo_email['Subject'] = "Aviso de Estoque Crítico"  # Assunto do e-mail
    conteudo_email['From'] = remetente_email
    conteudo_email['To'] = destinatario_email
 
    try:
        # Conexão segura com o servidor SMTP para envio do e-mail
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as servidor_smtp:
            servidor_smtp.login(remetente_email, senha_acesso)  # Login na conta
            servidor_smtp.send_message(conteudo_email)  # Envia a mensagem
        print(f"E-mail enviado: {mensagem_aviso}")
    except Exception as erro_envio:
        print(f"Erro ao enviar o e-mail: {erro_envio}")
 
 
def gerar_arquivo_excel(relatorio_dados):
    nome_relatorio = f"relatorio_estoque_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    relatorio_dados.to_excel(nome_relatorio, index=False)  # Salva o arquivo Excel
    print(f"Relatório gerado: {nome_relatorio}")
 
 
def processar_dados_estoque():
    print("Processamento dos dados de estoque iniciado...")
 
    # Tenta carregar os dados do estoque a partir de um arquivo CSV
    try:
        estoque_dados = pd.read_csv('Esp8266_Receiver.csv')
        print("Dados do estoque carregados com sucesso.")
    except FileNotFoundError:
        print("Erro: Arquivo 'Esp8266_Receiver.csv' não encontrado. Verifique o caminho.")
        return
 
    print(estoque_dados)  # Visualiza os dados carregados para referência
 
    contador_linhas = 0  # Contador para rastrear a linha sendo analisada
    lista_alertas = []  # Lista para armazenar mensagens de alerta
 
    # Percorre cada linha nas colunas de estoques das esteiras
    for indice, linha in enumerate(estoque_dados[['esteira1', 'esteira2', 'esteira3']].values):
        contador_linhas += 1
 
        # Verificação para cada esteira
        for esteira_id, valor in enumerate(linha, start=1):
            if valor == 1:
                aviso = f"Aviso: Estoque crítico na Esteira {esteira_id} (linha {contador_linhas})."
                print(aviso)
                lista_alertas.append(aviso)
                disparar_email_aviso(aviso)  # Dispara o aviso por e-mail
            elif valor == 2:
                print(f"Alerta: Estoque médio na Esteira {esteira_id} (linha {contador_linhas}). Planeje reabastecimento.")
            elif valor == 3:
                print(f"Status: Estoque cheio na Esteira {esteira_id} (linha {contador_linhas}). Sem ações necessárias.")
 
    # Adiciona informações de alertas ao DataFrame
    print("Adicionando informações de alertas ao relatório...")
    estoque_dados['Alerta Esteira 1'] = ["Crítico" if valores[0] == 1 else "" for valores in estoque_dados[['esteira1', 'esteira2', 'esteira3']].values]
    estoque_dados['Alerta Esteira 2'] = ["Crítico" if valores[1] == 1 else "" for valores in estoque_dados[['esteira1', 'esteira2', 'esteira3']].values]
    estoque_dados['Alerta Esteira 3'] = ["Crítico" if valores[2] == 1 else "" for valores in estoque_dados[['esteira1', 'esteira2', 'esteira3']].values]
 
    # Gera o relatório final em formato Excel
    gerar_arquivo_excel(estoque_dados)
 
    # Exibe um resumo dos alertas registrados
    if lista_alertas:
        print("Resumo dos alertas registrados:")
        for aviso in lista_alertas:
            print(aviso)
    else:
        print("Nenhum alerta crítico foi detectado.")
 
 
# Ponto de partida da execução do script
if __name__ == "__main__":
    processar_dados_estoque()