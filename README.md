# Automação de Indicadores (OnePage por loja)

Script em Python que transforma a planilha de vendas de uma rede de 25 lojas num relatório diário por e-mail: cada gerente recebe o resumo da própria loja, e a diretoria recebe o ranking de todas.

## O problema

Numa rede de lojas, alguém da equipe de dados abre a planilha de vendas toda manhã, separa os números de cada loja, calcula os indicadores, monta um resumo de uma página (o *OnePage*) e manda para cada gerente. São 25 e-mails feitos à mão, todo dia, com risco de enviar o dado de uma loja para o gerente errado.

## O que o script faz

1. **Lê as bases**: vendas (`Vendas.xlsx`), cadastro de lojas (`Lojas.csv`) e contatos dos gerentes (`Emails.xlsx`).
2. **Separa as vendas por loja** e salva um backup da planilha de cada uma em `Backup Arquivos Lojas/<loja>/`, com a data no nome do arquivo.
3. **Calcula os indicadores** do último dia disponível e do ano:

   | Indicador | Meta do dia | Meta do ano |
   |---|---|---|
   | Faturamento | R$ 1.000 | R$ 1.650.000 |
   | Diversidade de produtos | 4 | 120 |
   | Ticket médio | R$ 500 | R$ 500 |

4. **Envia o OnePage para cada gerente** pelo Outlook: tabela em HTML com valor, meta e um indicador verde/vermelho, e a planilha da loja em anexo. Cada gerente só recebe os dados da própria loja.
5. **Envia o ranking para a diretoria**: melhor e pior loja do dia e do ano no corpo do e-mail, com os rankings diário e anual em Excel anexados.

## Tecnologias

- **Python** + **pandas** (leitura, junção e agregação das bases)
- **pathlib** (criação das pastas de backup)
- **pywin32** (envio dos e-mails pelo Outlook)

## Como rodar

Precisa de Windows com o **Outlook instalado e logado**.

```bash
pip install pandas openpyxl pywin32
python Automacao.py
```

> Para testar, troque os e-mails dos gerentes em `Bases de Dados/Emails.xlsx` por um e-mail seu.

## Estrutura

```
Automacao.py                  script completo
Automacao de Processo.ipynb   versão em notebook, passo a passo
Descriçao do Projeto.ipynb    enunciado do desafio
Bases de Dados/               vendas, lojas e e-mails (dados fictícios)
Backup Arquivos Lojas/        saída: planilha de cada loja e rankings
```

---

Projeto de estudo de automação com Python · [Michaell Dantas](https://github.com/michaelldantas)
