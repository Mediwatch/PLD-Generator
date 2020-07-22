from docx import Document

def generate_chart(c, sort):

	c.add_paragraph("Projet : Mediwatch")
	for key in sort.keys():
		c.add_paragraph(f"Bloc {key}")
	c.add_page_break()