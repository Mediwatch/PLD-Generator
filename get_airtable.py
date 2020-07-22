import airtable
from strings import API_KEY

def get_airtable():
	cards = airtable.Airtable("appmxiObOjwHE5onb", "Cartes", API_KEY)
	advancements = airtable.Airtable("appmxiObOjwHE5onb", "Avancements", API_KEY)
	rapports = airtable.Airtable("appmxiObOjwHE5onb", "Rapport Personnel", API_KEY)
	return [cards.get_all(), advancements.get_all(), rapports.get_all()]