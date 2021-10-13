import pandas as pd
import sqlalchemy
import time

inicio = time.time()

engine = sqlalchemy.create_engine("mssql+pymssql://usuario_bd:senha_bd@ip_bd/nome_bd")

dados = pd.read_sql('''
        select par.codparc as 'Cód. Parceiro', par.nomeparc as 'Nome Parceiro', par.razaosocial as 'Identificação - Razão Social', par.email as 'Endereço - Email', par.emailnfe as 'NF-e/NFS-e/CT-e - E-mail p/ envio NF-e/NFS-e/CT-e', par.emailnfse as 'NF-e/NFS-e/CT-e - E-mail específico p/ envio NFS-e', par.emailnotifentrega as 'Endereço - E-mail específico p/ envio NFS-e', ctt.email as 'Contatos - Email', par.telefone as 'Endereço - Telefone', par.fax as 'Endereço - Celular/Fax', ctt.telefone as 'Contatos - Telefone', ctt.telresid as 'Contatos - Telefone Residencial', ctt.celular as 'Contatos - Celular'
        from tgfpar (nolock) par
        left join tgfctt (nolock) ctt on par.codparc = ctt.codparc
        where par.cliente = 'S' and par.codparc >= 50 and (par.email is not null or par.emailnfe is not null or par.emailnfse is not null or par.emailnotifentrega is not null or ctt.email is not null or par.telefone is not null or par.fax is not null or ctt.telefone is not null or ctt.telresid is not null or ctt.celular is not null)
        order by par.codparc asc
    ''', engine)

excel_writer = pd.ExcelWriter("contatos_de_clientes.xlsx", engine = "xlsxwriter")

dados.to_excel(excel_writer, sheet_name = "contatos_de_clientes")

excel_writer.save()

fim = time.time()

print(f"Processo demorou {fim - inicio} segundos")
