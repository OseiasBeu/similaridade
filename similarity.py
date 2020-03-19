import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

wb1 = pd.read_excel('herois1.xlsx')
wb2 = pd.read_excel('herois2.xlsx')

hr1 = pd.DataFrame(wb1)
hr2 = pd.DataFrame(wb2)

#print(hr1.head(20))
#print(hr2.head(20))

hr1['Hierarquia'] = hr1['Hierarquia'].apply(str)
hr2['Hierarquia'] = hr2['Hierarquia'].apply(str)

concact_hr1 = hr1.nome +hr1.Hierarquia +hr1.Cargo
concact_hr2 = hr2.nome +hr2.Hierarquia +hr2.Cargo
#print(concact_hr1.str.replace(' ',''))
#print(concact_hr2.str.replace(' ',''))

hr1['concate'] = concact_hr1
hr2['concate'] = concact_hr2

hr_full = pd.DataFrame()
hr_full['nome_hf1'] = hr1['nome']
hr_full['cargo_hf1'] = hr1['Cargo']
hr_full['hierarquia_hf1'] = hr1['Hierarquia']
hr_full['concate_hf1'] = hr1['concate']
hr_full['nome_hf2'] = hr2['nome']
hr_full['cargo_hf2'] = hr2['Cargo']
hr_full['hierarquia_hf2'] = hr2['Hierarquia']
hr_full['concate_hf2'] = hr2['concate']

percentual_ratio = []
percentual_partial = []
percentual_sort = []
for item_concatenado in range(len(hr1)):
    percentual_ratio.append(fuzz.ratio(hr1['concate'][item_concatenado],hr2['concate'][item_concatenado]))
    percentual_partial.append(fuzz.ratio(hr1['concate'][item_concatenado],hr2['concate'][item_concatenado]))
    percentual_sort.append(fuzz.token_sort_ratio(hr1['concate'][item_concatenado],hr2['concate'][item_concatenado]))
    #print(percentual)
    

percentualRatio_serie = pd.Series(percentual_ratio)
percentualPartial_serie = pd.Series(percentual_partial)
percentualSort_serie = pd.Series(percentual_sort)

#print(percentual_serie)
hr_full['percentual_comparacao_exata'] = percentualRatio_serie
hr_full['percentual_comparacao_aproximado'] = percentualPartial_serie
hr_full['percentual_comparacao_ordenacao'] = percentualSort_serie

print(hr_full)


try:
    hr_full.to_csv ('percentualDeSimilaridade.csv', index = False, header=True)
    print('Arquivo: {} geraldo com sucesso!'.format('percentualDeSimilaridade.csv'))
except ValueError as E:
    print(E)
    print('Falha na criação do arquivo!')