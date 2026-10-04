"""
Rimuove il placeholder {{ compenso_annuale }} frammentato dal template.
La riga resterà con solo "di " alla fine.
Il valore verrà iniettato programmaticamente da preventivi.pyw dopo il rendering.
"""
import zipfile, re, os

SRC = r'C:\Progetti\A-002 Gestione Compensi\modello_offerta.docx'
TMP = r'C:\Progetti\A-002 Gestione Compensi\_tmp_final.docx'

with zipfile.ZipFile(SRC, 'r') as zin:
    xml = zin.read('word/document.xml').decode('utf-8')

# Trova il paragrafo con "compenso annuale"
idx = xml.find('compenso annuale')
start_p = xml.rfind('<w:p ', 0, idx)
end_p   = xml.find('</w:p>', idx) + 6
para    = xml[start_p:end_p]

# Estrai il testo grezzo del paragrafo
testo = re.sub(r'<[^>]+>', '', para)
print('Testo attuale:', repr(testo))

# Trova il run che contiene "di " (l'ultimo run prima del placeholder)
# Dobbiamo tenere tutto fino al run "di " e rimuovere tutto il resto
# Il paragrafo deve finire con: ...di </w:t></w:r></w:p>

# Strategia: ricostruisci il paragrafo tenendo solo i run fino a "di "
# Trova l'ultimo run che contiene "di " nel raw
idx_di = para.rfind('>di <')
if idx_di >= 0:
    # Trova la fine di questo run: </w:r>
    end_di_run = para.find('</w:r>', idx_di) + 6
    # Tieni il paragrafo fino alla fine del run "di ", poi chiudi </w:p>
    pPr_match = re.search(r'<w:pPr>.*?</w:pPr>', para, re.DOTALL)
    pPr = pPr_match.group(0) if pPr_match else ''
    new_para = para[:end_di_run] + '</w:p>'
    print('Paragrafo rimosso placeholder OK')
else:
    print('Run "di " non trovato, provo con ">di<"')
    idx_di = para.rfind('>di<')
    if idx_di >= 0:
        end_di_run = para.find('</w:r>', idx_di) + 6
        new_para = para[:end_di_run] + '</w:p>'
        print('OK')
    else:
        print('ERRORE: run "di" non trovato')
        new_para = para  # lascia invariato

xml2 = xml[:start_p] + new_para + xml[end_p:]

# Verifica
testo2 = re.sub(r'<[^>]+>', '', xml2[start_p:start_p+len(new_para)])
print('Testo dopo fix:', repr(testo2))

# Salva
with zipfile.ZipFile(SRC, 'r') as zin:
    with zipfile.ZipFile(TMP, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == 'word/document.xml':
                zout.writestr(item, xml2.encode('utf-8'))
            else:
                zout.writestr(item, zin.read(item.filename))

os.replace(TMP, SRC)
print('Template salvato:', SRC)
