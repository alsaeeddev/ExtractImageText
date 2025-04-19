font_path = os.path.join(os.path.dirname(__file__), "DejaVuSans.ttf")
            pdf.add_font('DejaVu', '', font_path, uni=True)
            pdf.set_font('DejaVu', '', 12)