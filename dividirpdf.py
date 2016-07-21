import easygui
from PyPDF2 import PdfFileWriter, PdfFileReader 

path = easygui.fileopenbox()
savepath = easygui.diropenbox()
print savepath

infile = PdfFileReader(open(path, 'rb'), strict=False)

for i in xrange(infile.getNumPages()):
    p = infile.getPage(i)
    outfile = PdfFileWriter()
    outfile.addPage(p)
    with open(savepath + '\\' + 'pagina-%02d.pdf' % i, 'wb') as f:
        outfile.write(f)