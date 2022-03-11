import PyPDF2
import re
import pandas as pd
from openpyxl import Workbook

pdf = 'C:\\Users\\pablo.pamf\\Documents\\2021\\LerPDF\\RAIS declaracoes\\28 NEP\\EstabelecimentoCompleto18330836000110_2020.pdf'

dados = pd.DataFrame(columns=['cpf','nome','pis'])
pis_l = []
nome_l = []
cpf_l = []

with open(pdf, 'rb') as f:
    reader = PyPDF2.PdfFileReader(f)
    #PEGA CNPJ E ANO BASE
    pg = reader.getPage(0).extractText()
    pos_ano_base = pg.find("Ano Base:")
    pos_cnpj = pg.find("Para uso da empresa:")
    pos_razao = pg.find("Razão Social:")
    print(pg)
    #print(pos_ano_base, pos_cnpj, pos_razao)
    
    ano_base = pg[pos_ano_base+len("Ano Base:") : pos_ano_base+len("Ano Base:")+4]
    cnpj = pg[pos_cnpj+len("Para uso da empresa:"):pos_cnpj+len("Para uso da empresa:")+18]
    razao_social = pg[pos_razao+len("Razão Social:"): pos_razao+len("Razão social:")+50]
    razao_social_final = razao_social.find("Relatório")
    razao_social = razao_social[:razao_social_final]
    
    print(ano_base, cnpj, razao_social)


    ## LOCALIZA CPF, NOME e PIS em cada página
    for page in reader.pages:
        try:
            text = page.extractText()
            pos_pis = [m.start(0) for m in re.finditer("Para uso da empresa:", text)]
            for pis in pos_pis:
                    p = text[pis+len("Para uso da empresa:"):pis+len("Para uso da empresa:")+14]
                    #trata erro de busca de CNPJ
                    if p != cnpj[:-4]:
                        #print(p)
                        pis_l.append(p)
                    
            
            
            pos_nome = [m.start(0) for m in re.finditer("Para uso da empresa:", text)]
            
            for nome in pos_nome:
                    n = text[nome+len("Para uso da empresa:")+14:nome+len("Para uso da empresa:")+70]
                    n1 = n.find("Nascim")
                    n = n[:n1]
                    #Trata erro que busca CNPJ
                    if n[:4] != cnpj[-4:]:
                        #print(n)
                        nome_l.append(n)
                
            pos_cpf = [m.start(0) for m in re.finditer("Parcela Final", text)]
            for cpf in pos_cpf:
                    ate_cpf = len("Remun.12/06/19644 - Preta/negra0 - Nao deficienteM10 - Brasileiro-07 - Ensino médio completo.")
                    c = text[cpf+len("Parcela Final")+ate_cpf:cpf+len("Parcela Final")+ate_cpf+14]
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
    
    """ try:
        pos1 = total.find("CNPJ/CEI:")
        pos2 = total.find("Razão Social")
        pos3 = total.find("Data Abertura")
        if pos1>-1 and pos2 > -1 and pos3 > -1:
            print(pos1,pos2, pos3)
            print(total[pos1+len("CNPJ/CEI:") : pos2])
            print(total[pos2+len("Razão Social:") : pos3]) 
    except Exception:
        print("Erro ao localizar CNPJ") """
    

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