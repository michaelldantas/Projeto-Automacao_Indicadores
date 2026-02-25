import pandas as pd
import win32com.client as win32
import pathlib


#definição de metas
meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500


lojas = pd.read_csv(r"Bases de Dados\Lojas.csv", encoding="latin1", sep=";")
emails = pd.read_excel(r"Bases de Dados\Emails.xlsx")
vendas = pd.read_excel(r"Bases de Dados\Vendas.xlsx")

vendas = vendas.merge(lojas, on="ID Loja")
dicionario_lojas = {}
for loja in lojas["Loja"]:
    dicionario_lojas[loja] = vendas.loc[vendas["Loja"] == loja , :]

dia_indicador = vendas["Data"].max()



# Verificar se a pasta existe
caminho_backup = pathlib.Path("Backup Arquivos Lojas")
lista_arquivos_backup = [arquivo.name for arquivo in caminho_backup.iterdir()]

# Se não existir, cria a pasta com os nomes das lojas
for loja in dicionario_lojas:
    if loja not in lista_arquivos_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()

    local_arquivo = f"{dia_indicador.day}_{dia_indicador.year}_{loja}.xlsx"
    local_arquivo = caminho_backup / loja / local_arquivo
    dicionario_lojas[loja].to_excel(local_arquivo, index=False)

for loja in dicionario_lojas:
    
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja["Data"] == dia_indicador, :]

    # faturamento ano
    faturamento_ano = vendas_loja["Valor Final"].sum()

    # faturamento dia
    faturamento_dia = vendas_loja_dia["Valor Final"].sum()

    # diversidade de tickets
    qtd_prod_ano = len(vendas_loja["Produto"].unique())
    qtd_prod_dia = len(vendas_loja_dia["Produto"].unique())

    # tickt médio ano

    valor_venda_ano = vendas_loja.groupby("Código Venda").sum(numeric_only=True) 
    ticket_med_ano = valor_venda_ano["Valor Final"].mean()

    # tickt médio dia
    valor_venda_dia = vendas_loja_dia.groupby("Código Venda").sum(numeric_only=True)
    ticket_med_dia = valor_venda_dia["Valor Final"].mean()

    #enviar o e-mail para os gerentes
    outlook = win32.Dispatch('outlook.application')

    nome = emails.loc[emails['Loja']==loja, 'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja']==loja, 'E-mail'].values[0]
    mail.Subject = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'

    if faturamento_dia >= meta_faturamento_dia:
            cor_fat_dia = 'green'
    else:
            cor_fat_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
            cor_fat_ano = 'green'
    else:
            cor_fat_ano = 'red'
    if qtd_prod_dia >= meta_qtdeprodutos_dia:
            cor_qtde_dia = 'green'
    else:
            cor_qtde_dia = 'red'
    if qtd_prod_ano >= meta_qtdeprodutos_ano:
            cor_qtde_ano = 'green'
    else:
            cor_qtde_ano = 'red'
    if ticket_med_dia >= meta_ticketmedio_dia:
            cor_ticket_dia = 'green'
    else:
            cor_ticket_dia = 'red'
    if ticket_med_ano >= meta_ticketmedio_ano:
            cor_ticket_ano = 'green'
    else:
            cor_ticket_ano = 'red'

    mail.HTMLBody = f'''
        <p>Bom dia, {nome}</p>

        <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi:</p>

        <table>
        <tr>
            <th>Indicador</th>
            <th>Valor Dia</th>
            <th>Meta Dia</th>
            <th>Cenário Dia</th>
        </tr>
        <tr>
            <td>Faturamento</td>
            <td style="text-align: center">R${faturamento_dia:.2f}</td>
            <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
            <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
        </tr>
        <tr>
            <td>Diversidade de Produtos</td>
            <td style="text-align: center">{qtd_prod_dia}</td>
            <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
            <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
        </tr>
        <tr>
            <td>Ticket Médio</td>
            <td style="text-align: center">R${ticket_med_dia:.2f}</td>
            <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
            <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
        </tr>
        </table>
        <br>
        <table>
        <tr>
            <th>Indicador</th>
            <th>Valor Ano</th>
            <th>Meta Ano</th>
            <th>Cenário Ano</th>
        </tr>
        <tr>
            <td>Faturamento</td>
            <td style="text-align: center">R${faturamento_ano:.2f}</td>
            <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
            <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
        </tr>
        <tr>
            <td>Diversidade de Produtos</td>
            <td style="text-align: center">{qtd_prod_ano}</td>
            <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
            <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
        </tr>
        <tr>
            <td>Ticket Médio</td>
            <td style="text-align: center">R${ticket_med_ano:.2f}</td>
            <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
            <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
        </tr>
        </table>

        <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

        <p>Qualquer dúvida estou à disposição.</p>
        <p>Att., Michaell Dantas</p>
        '''




    attachment = pathlib.Path.cwd() / caminho_backup / loja / f"{dia_indicador.day}_{dia_indicador.year}_{loja}.xlsx"
    mail.Attachments.Add(str(attachment))
    mail.Send()
    print('E-mail da Loja {} enviado'.format(loja))

# faturamento do ano agrupadas por lojas
faturamento_lojas = vendas.groupby("Loja")[["Loja", "Valor Final"]].sum(numeric_only=True)
faturamento_lojas = faturamento_lojas.sort_values("Valor Final", ascending=False)

# Salvando o arquivo de faturamento do ano agrupadas por loja
local_arquivo = f"{dia_indicador.day}_{dia_indicador.month}_Ranking_Anual.xlsx"
local_arquivo = caminho_backup / local_arquivo
faturamento_lojas.to_excel(local_arquivo)

# faturamento do dia agrupadas por lojas
vendas_dia = vendas.loc[vendas["Data"] == dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby("Loja")[["Loja", "Valor Final"]].sum(numeric_only=True)
faturamento_lojas_dia = faturamento_lojas_dia.sort_values("Valor Final", ascending=False)

# Salvando o arquivo de faturamento do dia agrupadas por loja
local_arquivo = f"{dia_indicador.day}_{dia_indicador.month}_Ranking_Diario.xlsx"
local_arquivo = caminho_backup / local_arquivo
faturamento_lojas_dia.to_excel(local_arquivo)


#enviar o e-mail para a Diretoria
outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja']=="Diretoria", 'E-mail'].values[0]
mail.Subject = f'Ranking dia {dia_indicador.day}/{dia_indicador.month}'
mail.Body = f'''
Prezados, bom dia

Melhor loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0, 0]:,.2f}
Pior loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1, 0]:,.2f}

Melhor loja do Ano em Faturamento: Loja {faturamento_lojas.index[0]} com Faturamento R${faturamento_lojas.iloc[0, 0]:,.2f}
Pior loja do Ano em Faturamento: Loja {faturamento_lojas.index[-1]} com Faturamento R${faturamento_lojas.iloc[-1, 0]:,.2f}

Segue em anexo os rankings do ano e do dia de todas as lojas.

Qualquer dúvida estou à disposição.

Att.,
Michaell Dantas'''

attachment = pathlib.Path.cwd() / caminho_backup / f"{dia_indicador.day}_{dia_indicador.month}_Ranking_Anual.xlsx"
mail.Attachments.Add(str(attachment))
attachment = pathlib.Path.cwd() / caminho_backup / f"{dia_indicador.day}_{dia_indicador.month}_Ranking_Diario.xlsx"
mail.Attachments.Add(str(attachment))
mail.Send()
print('E-mail da diretoria enviado')