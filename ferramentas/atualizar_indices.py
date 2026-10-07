#!/usr/bin/env python3
"""Abre cada .docx no LibreOffice (headless), atualiza sumario e campos de pagina,
salva de volta em .docx e exporta um .pdf ao lado para conferencia.
Uso: python3 ferramentas/atualizar_indices.py arquivo1.docx [arquivo2.docx ...]"""
import os, sys, time, subprocess, uno
from com.sun.star.beans import PropertyValue

def prop(n, v):
    p = PropertyValue(); p.Name, p.Value = n, v; return p

proc = subprocess.Popen(['soffice', '--headless', '--invisible', '--norestore',
                         '--accept=socket,host=127.0.0.1,port=2083;urp;'],
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
ctx = None
for _ in range(60):
    try:
        local = uno.getComponentContext()
        res = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        ctx = res.resolve('uno:socket,host=127.0.0.1,port=2083;urp;StarOffice.ComponentContext'); break
    except Exception:
        time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
try:
    for f in sys.argv[1:]:
        url = uno.systemPathToFileUrl(os.path.abspath(f))
        doc = desktop.loadComponentFromURL(url, '_blank', 0, (prop('Hidden', True),))
        for _ in range(2):  # 2 passadas: o sumario muda a paginacao
            idx = doc.getDocumentIndexes()
            for i in range(idx.getCount()): idx.getByIndex(i).update()
            doc.getTextFields().refresh()
        doc.storeToURL(url, (prop('FilterName', 'MS Word 2007 XML'),))
        doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(f[:-5] + '.pdf')), (prop('FilterName', 'writer_pdf_Export'),))
        n = doc.getCurrentController().getPropertyValue('PageCount')
        print('OK', os.path.basename(f), n, 'paginas'); doc.close(True)
finally:
    try: desktop.terminate()
    except Exception: pass
    proc.wait(timeout=30)
