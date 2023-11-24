# -*- coding: utf-8 -*-
"""
Created on Fri Nov 24 10:43:54 2023

@author: suley
"""

import os 
import docx
import comtypes.client


word_path="Tuna_cover_letter.docx"
pdf_path="Tuna_cover_letter.pdf"

 
# Load the Word document using the docx library
doc = docx.Document(word_path)
 
# Save the Word document as a PDF using Microsoft Word
word = comtypes.client.CreateObject("Word.Application")
docx_path = os.path.abspath(word_path)
pdf_path = os.path.abspath(pdf_path)
 
 
 
pdf_format = 17  # PDF file format code
word.Visible = False
in_file = word.Documents.Open(docx_path)
in_file.SaveAs(pdf_path, FileFormat=pdf_format)
in_file.Close()
 
# Quit Microsoft Word
word.Quit()