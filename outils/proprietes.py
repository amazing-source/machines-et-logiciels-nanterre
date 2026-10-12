# -*- coding: utf-8 -*-
# Fixe l'auteur, le titre et le sujet d'un document Office (docx, xlsx, pptx)
# en reecrivant docProps/core.xml. Les proprietes integrees ne sont pas
# modifiables de facon fiable par COM depuis PowerShell (cf. HANDOFF 5.12).
# Usage : python proprietes.py <fichier> <auteur> <titre> <sujet>
import os
import re
import shutil
import sys
import tempfile
import zipfile
from xml.sax.saxutils import escape

path, author, title, subject = sys.argv[1:5]


def put(xml, tag, value):
    value = escape(value)
    pat = re.compile(r"<%s(\s[^>]*)?/>|<%s(\s[^>]*)?>.*?</%s>" % (tag, tag, tag), re.S)
    new = "<%s>%s</%s>" % (tag, value, tag)
    if pat.search(xml):
        return pat.sub(lambda m: new, xml, count=1)
    return xml.replace("</cp:coreProperties>", new + "</cp:coreProperties>")


fd, tmp = tempfile.mkstemp(suffix=os.path.splitext(path)[1])
os.close(fd)
with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "docProps/core.xml":
            xml = data.decode("utf-8")
            xml = put(xml, "dc:creator", author)
            xml = put(xml, "cp:lastModifiedBy", author)
            xml = put(xml, "dc:title", title)
            xml = put(xml, "dc:subject", subject)
            data = xml.encode("utf-8")
        # les medias deja compresses (ou non compressibles) restent tels quels
        ctype = zipfile.ZIP_STORED if item.compress_type == zipfile.ZIP_STORED else zipfile.ZIP_DEFLATED
        zout.writestr(item, data, compress_type=ctype)
shutil.move(tmp, path)
print("proprietes :", author, "|", title)
