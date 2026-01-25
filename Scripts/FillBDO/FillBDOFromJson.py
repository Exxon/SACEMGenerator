#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FillBDOFromJson.py
Remplit un PDF BDO SACEM avec des valeurs depuis un fichier JSON.
Appelé par BDOPdfGenerator.vb

Usage:
    python FillBDOFromJson.py <template.pdf> <values.json> <output.pdf>
"""

import sys
import json
from pypdf import PdfReader, PdfWriter


def fill_pdf(template_path: str, json_path: str, output_path: str):
    """
    Remplit le PDF template avec les valeurs du JSON.
    """
    # Lire les valeurs (utf-8-sig pour gérer le BOM)
    with open(json_path, 'r', encoding='utf-8-sig') as f:
        field_values = json.load(f)
    
    # Lire le template
    reader = PdfReader(template_path)
    writer = PdfWriter()
    writer.append(reader)
    
    # Remplir les champs sur chaque page
    for page in writer.pages:
        writer.update_page_form_field_values(page, field_values)
    
    # Écrire le PDF de sortie
    with open(output_path, 'wb') as f:
        writer.write(f)
    
    print(f"PDF rempli avec {len(field_values)} champs: {output_path}")


def main():
    if len(sys.argv) < 4:
        print("Usage: python FillBDOFromJson.py <template.pdf> <values.json> <output.pdf>")
        print("Exemple: python FillBDOFromJson.py Bdo711.pdf values.json BDO_rempli.pdf")
        sys.exit(1)
    
    template_path = sys.argv[1]
    json_path = sys.argv[2]
    output_path = sys.argv[3]
    
    fill_pdf(template_path, json_path, output_path)


if __name__ == "__main__":
    main()
