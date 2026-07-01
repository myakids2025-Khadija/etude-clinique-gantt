#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Génération d'un rapport PDF avec page de garde et capture du Gantt Chart.
"""

from fpdf import FPDF
import os
import shutil
from datetime import datetime

# ─── Chemins ───
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
IMAGE_SRC = r"C:\Users\THINKPAD\.gemini\antigravity\brain\b57a3f08-c7bc-42f3-8b87-54b633bb756a\media__1782631141613.png"
IMAGE_LOCAL = os.path.join(SCRIPT_DIR, "gantt_capture.png")
OUTPUT_PDF = os.path.join(SCRIPT_DIR, "Rapport_Planning_Realisation.pdf")
ARIAL_TTF = r"C:\WINDOWS\Fonts\arial.ttf"
ARIALB_TTF = r"C:\WINDOWS\Fonts\arialbd.ttf"
ARIALI_TTF = r"C:\WINDOWS\Fonts\ariali.ttf"

# Copier l'image localement
shutil.copy2(IMAGE_SRC, IMAGE_LOCAL)
print(f"[OK] Image copiée : {IMAGE_LOCAL}")


class RapportPDF(FPDF):
    def __init__(self):
        super().__init__(orientation='P', unit='mm', format='A4')
        # Polices Unicode
        self.add_font('Arial', '', ARIAL_TTF, uni=True)
        self.add_font('Arial', 'B', ARIALB_TTF, uni=True)
        self.add_font('Arial', 'I', ARIALI_TTF, uni=True)
        self.page_count = 0

    def header(self):
        if self.page_no() > 1:
            self.set_font('Arial', 'I', 8)
            self.set_text_color(120, 120, 120)
            self.cell(0, 8, 'Rapport — Planning de Réalisation | Pôle Médical', 0, 0, 'L')
            self.cell(0, 8, f'Page {self.page_no() - 1}', 0, 1, 'R')
            self.set_draw_color(59, 130, 246)
            self.set_line_width(0.5)
            self.line(10, 14, 200, 14)
            self.ln(6)

    def footer(self):
        if self.page_no() > 1:
            self.set_y(-15)
            self.set_font('Arial', 'I', 7)
            self.set_text_color(150, 150, 150)
            self.cell(0, 10, f'© {datetime.now().year} — Document confidentiel', 0, 0, 'C')

    # ─── Page de garde ───
    def page_de_garde(self):
        self.add_page()

        # Bande bleue en haut
        self.set_fill_color(15, 23, 42)   # Bleu foncé
        self.rect(0, 0, 210, 100, 'F')

        # Bande accent
        self.set_fill_color(59, 130, 246)  # Bleu accent
        self.rect(0, 95, 210, 6, 'F')

        # Logo / Icône
        self.set_y(25)
        self.set_font('Arial', 'B', 42)
        self.set_text_color(59, 130, 246)
        self.cell(0, 15, '⬛', 0, 1, 'C')

        # Titre principal
        self.set_y(48)
        self.set_font('Arial', 'B', 28)
        self.set_text_color(255, 255, 255)
        self.cell(0, 14, 'PLANNING DE RÉALISATION', 0, 1, 'C')

        # Sous-titre
        self.set_font('Arial', '', 14)
        self.set_text_color(148, 163, 184)
        self.cell(0, 10, 'Réalisation des travaux d\'un pôle médical', 0, 1, 'C')

        # Ligne décorative
        self.ln(5)
        self.set_draw_color(59, 130, 246)
        self.set_line_width(0.8)
        self.line(70, self.get_y(), 140, self.get_y())

        # Informations du projet
        self.ln(20)
        self.set_text_color(30, 41, 59)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'INFORMATIONS DU PROJET', 0, 1, 'C')
        self.ln(5)

        # Tableau d'informations
        infos = [
            ('Objet', 'Planning de réalisation — Pôle médical'),
            ('Délai d\'exécution', 'À partir de l\'ODS notifié au service'),
            ('', 'co-contractant par le service contractant'),
            ('Version 1', 'Délai de 24 mois'),
            ('Version 2', 'Délai optimisé de 16 mois (×0.667)'),
            ('Date du document', datetime.now().strftime('%d/%m/%Y')),
        ]

        self.set_font('Arial', '', 10)
        for label, value in infos:
            self.set_x(35)
            if label:
                self.set_font('Arial', 'B', 10)
                self.set_text_color(59, 130, 246)
                self.cell(50, 8, label, 0, 0, 'R')
                self.set_text_color(51, 65, 85)
                self.set_font('Arial', '', 10)
                self.cell(5, 8, '  :  ', 0, 0, 'C')
                self.cell(0, 8, value, 0, 1, 'L')
            else:
                self.cell(55, 8, '', 0, 0)
                self.set_text_color(51, 65, 85)
                self.set_font('Arial', '', 10)
                self.cell(0, 8, value, 0, 1, 'L')

        # Badge version
        self.ln(15)
        self.set_fill_color(240, 245, 255)
        self.set_draw_color(59, 130, 246)
        self.set_line_width(0.3)
        self.set_x(45)
        self.rect(45, self.get_y(), 120, 20, 'FD')
        self.set_font('Arial', 'B', 10)
        self.set_text_color(59, 130, 246)
        self.cell(0, 20, 'DOCUMENT CONFIDENTIEL — USAGE INTERNE', 0, 1, 'C')

        # Pied de page de garde
        self.set_y(-35)
        self.set_draw_color(200, 200, 200)
        self.set_line_width(0.2)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(5)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 6, f'Généré le {datetime.now().strftime("%d/%m/%Y à %H:%M")}', 0, 1, 'C')

    # ─── Page du Gantt Chart ───
    def page_gantt(self):
        self.add_page()

        # Titre de section
        self.set_font('Arial', 'B', 18)
        self.set_text_color(15, 23, 42)
        self.cell(0, 12, 'Diagramme de Gantt — Version 16 Mois', 0, 1, 'L')

        # Sous-titre
        self.set_font('Arial', '', 10)
        self.set_text_color(100, 116, 139)
        self.multi_cell(0, 6,
            'Le diagramme ci-dessous présente le planning de réalisation recalculé sur 16 mois '
            '(facteur de compression ×0.667 appliqué à toutes les durées et dates de début '
            'par rapport à la version originale de 24 mois).')
        self.ln(5)

        # Barre d'info
        self.set_fill_color(240, 248, 255)
        self.set_draw_color(59, 130, 246)
        self.set_line_width(0.3)
        y_box = self.get_y()
        self.rect(10, y_box, 190, 14, 'FD')
        self.set_xy(12, y_box + 2)
        self.set_font('Arial', 'B', 9)
        self.set_text_color(59, 130, 246)
        self.cell(40, 10, 'Délai d\'exécution :', 0, 0)
        self.set_font('Arial', '', 9)
        self.set_text_color(51, 65, 85)
        self.cell(0, 10, 'À partir de l\'ODS notifié au service co-contractant par le service contractant.', 0, 1)
        self.ln(8)

        # Image du Gantt (pleine largeur)
        img_y = self.get_y()
        self.set_draw_color(42, 53, 80)
        self.set_line_width(0.4)
        # Image en paysage sur la largeur
        img_w = 190
        self.image(IMAGE_LOCAL, x=10, y=img_y, w=img_w)
        # Cadre autour
        # On calcule la hauteur proportionnelle
        from PIL import Image
        with Image.open(IMAGE_LOCAL) as img:
            w_px, h_px = img.size
        img_h = img_w * h_px / w_px
        self.rect(10, img_y, img_w, img_h, 'D')

        self.set_y(img_y + img_h + 8)

        # Légende
        self.set_font('Arial', 'I', 8)
        self.set_text_color(120, 120, 120)
        self.cell(0, 6, 'Figure 1 — Diagramme de Gantt interactif, version 16 mois (capture écran)', 0, 1, 'C')

    # ─── Page récapitulatif des tâches ───
    def page_recapitulatif(self):
        self.add_page()

        self.set_font('Arial', 'B', 18)
        self.set_text_color(15, 23, 42)
        self.cell(0, 12, 'Récapitulatif des tâches', 0, 1, 'L')

        self.set_font('Arial', '', 10)
        self.set_text_color(100, 116, 139)
        self.cell(0, 8, 'Comparaison des durées entre la version 24 mois et la version optimisée 16 mois.', 0, 1)
        self.ln(6)

        # En-têtes du tableau
        col_widths = [8, 65, 30, 12, 30, 25]
        headers = ['#', 'Tâche', 'V1 (24 mois)', '', 'V2 (16 mois)', 'Début V2']

        self.set_fill_color(15, 23, 42)
        self.set_text_color(255, 255, 255)
        self.set_font('Arial', 'B', 8)
        for i, (w, h) in enumerate(zip(col_widths, headers)):
            self.cell(w, 9, h, 1, 0, 'C', True)
        self.ln()

        # Données
        tasks = [
            ('1', 'Installation de chantier', '15 jours', '→', '10 jours', 'M0'),
            ('2', 'Terrassements / Excavations', '3 mois', '→', '2 mois', '10J'),
            ('3', 'Infrastructure', '7 mois', '→', '5 mois', 'M2'),
            ('4', 'Superstructure', '12 mois', '→', '8 mois', 'M5'),
            ('5', 'Maçonnerie et enduits', '12 mois', '→', '8 mois', 'M5'),
            ('6', 'VRD', '17 mois', '→', '11 mois', 'M2½'),
            ('7', 'Électricité / Plomberie', '17 mois', '→', '11 mois', 'M3'),
            ('8', 'Clim / Chauffage central', '16 mois', '→', '11 mois', 'M3'),
            ('9', 'Finition sol / Peinture', '13 mois', '→', '9 mois', 'M6'),
            ('10', 'Nettoyage', '2 mois', '→', '1 mois ½', 'M14'),
            ('11', 'Essais + Formation', '2 mois ½', '→', '1 mois ½', 'M14'),
            ('12', '◆ Réception provisoire', 'Mois 24', '→', 'Mois 16', 'M16'),
        ]

        self.set_font('Arial', '', 9)
        fill = False
        for row in tasks:
            if fill:
                self.set_fill_color(245, 247, 250)
            else:
                self.set_fill_color(255, 255, 255)

            self.set_text_color(51, 65, 85)
            self.cell(col_widths[0], 8, row[0], 1, 0, 'C', True)

            self.set_font('Arial', 'B', 9)
            self.set_text_color(15, 23, 42)
            self.cell(col_widths[1], 8, row[1], 1, 0, 'L', True)

            self.set_font('Arial', '', 9)
            self.set_text_color(59, 130, 246)
            self.cell(col_widths[2], 8, row[2], 1, 0, 'C', True)

            self.set_text_color(245, 158, 11)
            self.cell(col_widths[3], 8, row[3], 1, 0, 'C', True)

            self.set_text_color(249, 115, 22)
            self.set_font('Arial', 'B', 9)
            self.cell(col_widths[4], 8, row[4], 1, 0, 'C', True)

            self.set_font('Arial', '', 9)
            self.set_text_color(100, 116, 139)
            self.cell(col_widths[5], 8, row[5], 1, 1, 'C', True)
            fill = not fill

        self.ln(10)

        # Note de bas
        self.set_fill_color(255, 251, 235)
        self.set_draw_color(245, 158, 11)
        self.set_line_width(0.3)
        y = self.get_y()
        self.rect(10, y, 190, 18, 'FD')
        self.set_xy(14, y + 2)
        self.set_font('Arial', 'B', 9)
        self.set_text_color(180, 83, 9)
        self.cell(0, 6, 'Note :', 0, 1)
        self.set_x(14)
        self.set_font('Arial', '', 8)
        self.set_text_color(120, 80, 20)
        self.cell(0, 6, 'Le facteur de compression ×0.667 (16/24) a été appliqué uniformément à toutes les durées et dates de début.', 0, 1)


def main():
    pdf = RapportPDF()
    pdf.set_auto_page_break(auto=True, margin=20)

    print("[...] Génération de la page de garde...")
    pdf.page_de_garde()

    print("[...] Ajout du diagramme de Gantt...")
    pdf.page_gantt()

    print("[...] Ajout du récapitulatif...")
    pdf.page_recapitulatif()

    pdf.output(OUTPUT_PDF)
    print(f"\n[OK] Rapport PDF généré : {OUTPUT_PDF}")


if __name__ == '__main__':
    main()
