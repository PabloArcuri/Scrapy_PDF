import PyPDF2
import re
import pandas as pd
from openpyxl import Workbook

pdf = 'C:\\Users\\pablo.pamf\\Documents\\2021\\LerPDF\\RAIS declaracoes\\17 Chique\\EstabelecimentoCompleto16684906000102_2016 1.pdf'

dados = pd.DataFrame(columns=['cpf','nome','pis'])
pis_l = []
nome_l = []
cpf_l = []

with open(pdf, 'rb') as f:
    reader = PyPDF2.PdfFileReader(f)
    #PEGA CNPJ E ANO BASE
    pg = reader.getPage(0).extractText()
    pos_ano_base = pg.find("Ano-Base:")
    pos_cnpj = pg.find("CNPJ/CEI:")
    pos_razao = pg.find("Razão Social:")
    ##print(pg)
    
    ano_base = pg[pos_ano_base+10 : pos_ano_base+14]
    cnpj = pg[pos_cnpj+len("CNPJ/CEI:"):pos_razao]
    razao_social = pg[pos_razao+len("Razão Social:"): pos_razao+len("Razão social:")+50]
    razao_social_final = razao_social.find("Data Abertura")
    razao_social = razao_social[:razao_social_final]
    
    print(ano_base, cnpj, razao_social)


    ## LOCALIZA CPF, NOME e PIS em cada página
    for page in reader.pages:
        try:
            text = page.extractText()
            pos_pis = [m.start(0) for m in re.finditer("PIS/PASEP:", text)]
            for pis in pos_pis:
                    p = text[pis+len("PIS/PASEP:"):pis+24]
                    #print(p)
                    pis_l.append(p)
                    
            
            
            pos_nome = [m.start(0) for m in re.finditer("Nome:", text)]
            for nome in pos_nome:
                    n = text[nome+len("Nome:"):nome+50]
                    n1 = text[nome+len("Nome:"):nome+50].find("CPF")
                    n = n[:n1]
                    #print(n)
                    nome_l.append(n)
            
            pos_cpf = [m.start(0) for m in re.finditer("Nacionalidade:", text)]
            for cpf in pos_cpf:
                    c = text[cpf+len("Nacionalidade:"):cpf+len("Nacionalidade:")+14]
                    #print(c)
                    cpf_l.append(c)
            
            #print(c)
            
            #pos_pis.append(pos_pis)
            #print(pos_pis, reader.getPageNumber(page))
            
            #print(text[pos_pis[page]:pos_pis[page]+11], reader.getPageNumber(page))
            #pos1 = text.find("CNPJ")
            #print(pos1, reader.getPageNumber(page))
        except ValueError:
            print("Não localizado")

    #print(total)
    ## LOCALIZA CPNPJ
    
     

dados['cpf'] = cpf_l
dados['pis'] = pis_l
dados['nome'] = nome_l
dados['cnpj'] = cnpj
dados['ano_base'] = ano_base
print(dados)
book = Workbook()
sheet = book.active
book.save(razao_social+' - '+ano_base+'.xlsx')
try:
    cria_excel_total = pd.ExcelWriter('./'+razao_social+' - '+ano_base+'.xlsx')
    dados.to_excel(cria_excel_total, sheet_name= 'Total', index=False) 
    cria_excel_total.close()
except ValueError:
            print("Erro ao gerar excel")