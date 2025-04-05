import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

# Synonym dictionary
synonym_map = {
    # Office Supplies
    "sticky note": "post-it",
    "post it": "sticky note",
    "post-it": "sticky note",
    "scotch tape": "tape",
    "clear tape": "tape",
    "packing tape": "tape",
    "duct tape": "tape",
    "sharpie": "permanent marker",
    "permanent marker": "marker",
    "highlighter pen": "marker",
    "dry erase marker": "whiteboard marker",
    "notebook": "paper",
    "legal pad": "paper",
    "printer paper": "paper",
    "copy paper": "paper",
    "writing tool": "pen",
    "gel pen": "pen",
    "ballpoint": "pen",
    "mechanical pencil": "pencil",
    "staple gun": "stapler",
    "staples": "stapler",
    "clip": "binder clip",
    "paper clip": "clip",
    "binder clamp": "binder clip",
    "white-out": "correction fluid",
    "correction tape": "correction fluid",
    "eraser": "correction fluid",
    "file folder": "folder",
    "hanging folder": "folder",
    "accordion folder": "folder",
    "manila folder": "folder",
    "rubber bands": "rubber band",
    "push pins": "thumbtack",
    "tack": "thumbtack",
    "sticky tack": "thumbtack",

    # Snacks & Beverages
    "coffee pod": "coffee",
    "keurig pod": "coffee",
    "decaf": "coffee",
    "tea bag": "tea",
    "herbal": "tea",
    "green tea": "tea",
    "bottled water": "water",
    "spring water": "water",
    "sparkling water": "water",
    "pop": "soda",
    "coke": "soda",
    "pepsi": "soda",
    "cola": "soda",
    "can drink": "soda",
    "snack bar": "granola bar",
    "energy bar": "granola bar",
    "chips": "snacks",
    "pretzels": "snacks",
    "crackers": "snacks",
    "cookies": "snacks",
    "instant noodles": "ramen",
    "cup noodles": "ramen",

    # Mailroom Supplies
    "bubble mailer": "envelope",
    "padded envelope": "envelope",
    "manila envelope": "envelope",
    "stamps": "postage",
    "shipping label": "label",
    "return label": "label",
    "box cutter": "utility knife",
    "blade": "utility knife",
    "cutting knife": "utility knife",
    "packing peanuts": "packing material",
    "air pouch": "packing material",
    "bubble wrap": "packing material"
}

# Load the Excel file
def load_data():
    file_path = "Office Supply Inventory Project_v1.xlsx"
    df = pd.read_excel(file_path, sheet_name='v1')
    df = df.dropna(subset=['Item name'])
    return df

def fuzzy_search(df, query, threshold=60):
    # Perform fuzzy matching on 'Item name'
    matches = []
    for index, row in df.iterrows():
        score = fuzz.token_set_ratio(query.lower(), str(row['Item name']).lower())
        if score >= threshold:
            matches.append((score, row))
    matches.sort(reverse=True, key=lambda x: x[0])
    return [match[1] for match in matches]

# Streamlit app
st.title("📂 Office Supply Search")

query = st.text_input("What supply are you looking for? (e.g. pen, sticky note, binder)")

if query:
    df = load_data()

    # Try synonym first
    normalized_query = synonym_map.get(query.lower(), query.lower())
    results = fuzzy_search(df, normalized_query)

    # If no results, try original query as fallback
    if not results and normalized_query != query.lower():
        results = fuzzy_search(df, query.lower())

    if results:
        st.success(f"Found {len(results)} matching item(s):")
        for item in results:
            st.markdown(f"**{item['Item name']}**")
            st.write(f"Room: {item['Room']}, Shelf/Cabinet: {item['Shelf/Cabinet']}, Status: {item['Status']}")
            st.write("---")
    else:
        st.warning("No matching items found. Try another keyword.")
else:
    st.info("Enter a keyword to begin your search.")
