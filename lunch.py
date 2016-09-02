#from openpyxl import load_workbook
from __future__ import print_function
#import warnings
#~ warnings.filterwarnings("ignore")

#~ ficheiroFonte = ""

#Files with source data
#~ wbFonte = load_workbook('F:\\Work\\Ementas\\ementa7_20_07_2016.xlsx')
#~ wsFonte = wbFonte.worksheets[0]

#~ #Target Files
#~ wbTarget = load_workbook('F:\\Work\\Ementas\\teste.xlsx')
#~ wsTarget = wbTarget.worksheets[0]

#MODDED FONTE IN A list of lists
salasFonte = [[5,6,8,9,10],
	[12,13,14,15,16,17],
	[19,20,21,22,23,24],
	[26,27,28,29,30,31],
	[33,34,35,36]]
letrasF = 	['B','E','H','K','N']
	
#FONTE
#s1BerF = 	
#s2BerF = 	
#sParqueF = 	
#sPreCatlF = 
#dietaF = 	



#TARGET
salasTarget = [[6,7,9,10,11],
	[15,16,17,18,19,20],
	[24,25,26,27,28,29],
	[33,34,35,36,37,38],
	[42,43,44,45]]
letrasT = 	['B','C','D','E','F']

#~ s1BerT = 	
#~ s2BerT = 	
#~ sParqueT = 	
#~ sPreCatlT =     
#~ dietaT = 	
#salas = 	[s1BerT,s2Ber,sParque,sPreCatl,dieta]

#LOGOTIPO
#logo = "F:\\Raul\Work\Ementas\Logotipo_alta_png.png"


for lF in range(len(letrasF)):
	print("\n	",letrasF[lF])
	for nrS in range(len(salasFonte)):
		print("\n Sala: ",nrS)
		for nrL in range(len(salasFonte[nrS])):
			print(nrL,end=' ')


#~ #SALA BER1
#~ for i in range(len(letrasF)):
	#~ for k in range(len(s1BerF)):
			#~ wsTarget["%s%d" % (letrasT[i],s1BerT[k])].value = wsFonte["%s%d" % (letrasF[i],s1BerF[k])].value
				
#~ #SALA BER2
#~ for i in range(len(letrasF)):
	#~ for k in range(len(s2BerF)):
			#~ wsTarget["%s%d" % (letrasT[i],s2BerT[k])].value = wsFonte["%s%d" % (letrasF[i],s2BerF[k])].value

#~ #SALA PARQUE
#~ for i in range(len(letrasF)):
	#~ for k in range(len(sParqueF)):
			#~ wsTarget["%s%d" % (letrasT[i],sParqueT[k])].value = wsFonte["%s%d" % (letrasF[i],sParqueF[k])].value

#~ #PRE/CATL
#~ for i in range(len(letrasF)):
	#~ for k in range(len(sPreCatlF)):
			#~ wsTarget["%s%d" % (letrasT[i],sPreCatlT[k])].value = wsFonte["%s%d" % (letrasF[i],sPreCatlF[k])].value
#~ #DIETA
#~ for i in range(len(letrasF)):
	#~ for k in range(len(dietaF)):
			#~ wsTarget["%s%d" % (letrasT[i],dietaT[k])].value = wsFonte["%s%d" % (letrasF[i],dietaF[k])].value
		
#wbTarget.save('teste7_20_07_2016.xlsx')
