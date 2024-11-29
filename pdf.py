from spire.doc import *
from spire.doc.common import *

def convertToPdf(file, folder):
    documentpdf = Document()
    documentpdf.LoadFromFile(name)
    documentpdf.Watermark = None
    documentpdf.SaveToFile(folder+".pdf", FileFormat.PDF)
    documentpdf.Close()