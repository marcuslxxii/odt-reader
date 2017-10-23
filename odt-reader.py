#!/usr/bin/env python3
"""
Modulo che si propone di restituire, dato un file odt (basato su
LibreOffice v 5.4.2.2), il solo testo contenuto, senza usare moduli
aggiuntivi ma solo quello che mette a disposizione una installazione
standard di Python 3. In particolare si usa solo zipfile.
Si può usare questo modulo per leggere una serie di file odt (magari
centinaia) per analizzare i dati contenuti e poter generare un report.
Non si sarà quindi costretti ad installare un nuovo modulo, ma basta
copiare questo file accando al proprio, in cui lo si importa con:
from odt-reader import odt2str
e lo si richiama con:
testo = odt2str(nome_file)
oppure:
testo = odt2str(nome_file, False)
se si vogliono mantenere gli spazi speciali ('\xa0', creati su Writer
con Ctrl+Shift+Space) e quindi non sostituirli con spazi normali, o
anche:
testo = odt2str(nome_file, True, '\r\n')
se si vuole che l'andata a capo sia alla "maniera di Windows" ('\r' è
alla "maniera di Mac"), altrimenti il default è '\n' che è "alla
maniera di Linux". In fase di analisi, bisogna sapere qual è questo
carattere (indipendentemente dal SO). Quando serve specificarlo? Solo se
si vuol salvare in un file di testo il contenuto di un file odt, allora
bisogna specificare il carattere dell'andata a capo del SO in cui si
aprirà quel file di testo.

Si può testare questo modulo richiamandolo dalla shell di Linux o dal
prompt dei comandi di Windows e ponendo come parametri il percorso ed
i nomi dei file odt da leggere. In questo caso verranno importati anche
sys e os (entrambi già presenti, come zipfile, nell'installazione
standard di Python 3). Es.: odtreader.py esempio.odt

ATTENZIONE: non fa differenza tra Invio e Shift+Invio, non supporta le
tabelle, né elementi grafici diversi dai checkbox (riportati con "[ ]"
quelli deselezionati e con "[X]" quelli selezionati) e radiobutton
(riportati con "( )" quelli deselezionati e con "(X)" quelli
selezionati).

   Copyright ©2017 Marco Garipoli
   
   This program is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   This program is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   You should have received a copy of the GNU General Public License
   along with this program.  If not, see <http://www.gnu.org/licenses/>.
"""

# Lunghezza tot: 79 caratteri
# 34567890123456789012345678901234567890123456789012345678901234567890123456789

# Lunghezza delle Docstring: 72 caratteri
# 3456789012345678901234567890123456789012345678901234567890123456789012


__author__ = 'Marco Garipoli marcuslxxii@gmail.com'


__version__ = "$Revision: 000000000001 $"
# $Source$


def odt2str(nome_file, normalizza_spazi=True, a_capo='\n') -> str:
    """
    Legge il file odt passato (completo di path) e lo restituisce come
    testo semplice.
    In particolare per andare a capo sarà usato il/i carattere/i
    specificato/i nella chiamata (di default '\n').
    Qualunque segno grafico sarà totalemente ignorato, mentre invece le
    tabulazioni saranno riportate fedelmente, così come le righe vuote.
    v1.0 (23/10/2017)
    
    :type nome_file: str
    :param nome_file: percorso e nome del file odt da leggere
    
    :type a_capo: str
    :param a_capo: carattere/i da usare per l'andata a capo dell'output
    
    :type normalizza_spazi: bool
    :param normalizza_spazi: se True, sostituisce gli spazi speciali
                             (Ctrl+Shift+Space) con spazi normali
    """
    
    import zipfile as zf
    
    def pulisci_span(testo) -> str:
        """
        Se ci sono sottostrighe "<text:span...>xxx</text:span>" lascia
        solo "xxx"
        :type testo: str
        """
        nonlocal forms
        sub_out = ''
        i_t = testo.find('<text:')
        i_d = testo.find('<draw:')
        if i_t > -1 and (i_d == -1 or i_t < i_d):
            sub_out = testo[:i_t]
            if testo[i_t+6:].startswith('tab/>'):
                sub_out += '\t' + pulisci_span(testo[i_t+11:])
            elif testo[i_t+6:].startswith('span '):
                i = testo.find('>', i_t+11) + 1
                if testo[i-2] != '/':
                    j = testo.find('</text:span>', i)
                    sub_out += (pulisci_span(testo[i:j]) +
                                pulisci_span(testo[j+12:]))
            #else:
            #    print('\nSto ignorando: "{}"\n'.format(testo[i_t:]))
        elif i_d > -1 and (i_t == -1 or i_d < i_t):
            j = testo.find('/>', i_d+6)
            # Tra i_d e j c'è l'oggetto grafico. Ne rappresento qualcuno.
            rappr = ''
            t = testo[i_d+6:j]  # Testo del draw tra "<draw:" e "/>"
            i2 = t.find('draw:control="', i_d)
            if t.startswith('control ') and i2 > 0:
                j2 = t.find('"', i2+14)
                nome_controllo = t[i2+14:j2]  # Virgolette escluse
                #print('Trovato: ' + nome_controllo)
                if nome_controllo in forms.keys():
                    rappr = forms[nome_controllo]
            sub_out = testo[:i_d] + rappr + pulisci_span(testo[j+2:])
        else:
            sub_out = testo
        return sub_out
    
    forms = {}  # Nomi e rappresentaz di eventuali checkbox e radiobutton
    if not nome_file.lower().endswith('.odt'):
        nome_file += '.odt'
    out = ''
    with zf.ZipFile(nome_file) as z:
        with z.open('content.xml') as f:
            t = f.read().decode("utf-8") 
    if t != '':
        #t = t.replace('<text:tab/>', '\t')  # Va fatto prima dei find!
        i = t.find('<form:form')
        if i > -1:
            j = t.find('</form:form>')
            forms_info = t[i:j]  # Info su eventuali checkbox e radiobutton
            i2 = forms_info.find('<form:')
            while i2 > -1:
                j2 = forms_info.find('</form:', i2)  # Fine tag
                nome_controllo = ''
                if forms_info[i2+6:].startswith('checkbox '):
                    # Mi serve a questo punto sapere innanzitutto l'ID...
                    k1 = forms_info.find('xml:id="', i2)  # 'form:id="'
                    k2 = forms_info.find('"', k1+8)       # k1+9
                    nome_controllo = forms_info[k1+8:k2]  # k1+9
                    
                    # ...e poi lo stato (se selezionato o deselezionato)
                    r = '[ ]'
                    k1 = forms_info.find('form:current-state="', i2)
                    if 0 < k1 < j2:
                        k2 = forms_info.find('"', k1+20)
                        stato = forms_info[k1+20:k2]  # checked
                        if stato == "checked":
                            r = '[X]'
                elif forms_info[i2+6:].startswith('radio '):
                    # Mi serve a questo punto sapere innanzitutto l'ID...
                    k1 = forms_info.find('xml:id="', i2)  # 'form:id="'
                    k2 = forms_info.find('"', k1+8)       # k1+9
                    nome_controllo = forms_info[k1+8:k2]  # k1+9
                    
                    # ...e poi lo stato (se selezionato o deselezionato)
                    r = '( )'
                    k1 = forms_info.find('form:current-selected="', i2)
                    if 0 < k1 < j2:
                        k2 = forms_info.find('"', k1+23)
                        stato = forms_info[k1+23:k2]  # checked
                        if stato == "true":
                            r = '(X)'
                if nome_controllo != '' and not nome_controllo in forms.keys():
                    forms[nome_controllo] = r
                
                i2 = forms_info.find('<form:', j2)
            
            i = j+12
        #print('\nforms = {}\n'.format(forms))
        i = t.find('<text:p', i+1)
        while i > -1:
            i = t.find('>', i+7) + 1  # inizio testo
            if t[i-2] != '/':
                j = t.find('</text:p>', i)  # fine testo
                # Il testo interessato in questo passaggio sta in t[i:j]
                out += pulisci_span(t[i:j]) + a_capo  # '\n'
                i = t.find('<text:p', j+7)
            else:
                out += a_capo  # '\n'
                i = t.find('<text:p', i)
    out = out.replace('&lt;', '<')
    out = out.replace('&gt;', '>')
    if normalizza_spazi:
        out = out.replace('\xa0', ' ')  # '\xa0'=chr(160): spazio "appiccicato"
                                        # ' '   =chr( 32): spazio normale
    return out


def main():
    import sys, os
    
    num_param = len(sys.argv)  # sys.argv[0] è il nome del file .py
    if num_param > 1:
        a_capo = os.linesep  # Usa il separatore del S.O.
        particolareggiata = False
        for i in list(range(1, num_param))[:]:
            if sys.argv[i] == '-p':
                particolareggiata = True
                num_param -= 1
            else:
                if num_param > 2:
                    if i > 1:
                        print('='*79)
                    print('{}) {}'.format(i, sys.argv[i], True, a_capo))
                    print('-'*79)
                if particolareggiata:
                    particolareggiata = False  # -p si specifica per ogni file
                    t = odt2str(sys.argv[i], False)
                    j = 1
                    for c in t:
                        if c != '\n':
                            print('"{}" -> ascii({})'.format(c, ord(c)),
                                                             end='\t')
                            if j%3 == 0:
                                print()
                        else:
                            j = 0
                            print('"[Invio]" -> ascii({})'.format(ord(c)))
                        j += 1
                    print()
                else:
                    print(odt2str(sys.argv[i], True, a_capo))
    else:
        print("Indicare il/i file odt leggere.")
        print("Usare l'opzione -p per avere una stampa particolareggiata.")


if __name__ == "__main__":
    main()

