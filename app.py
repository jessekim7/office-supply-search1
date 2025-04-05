import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

# Synonym dictionary
synonym_map = {
    "sticky note": "post-it",
    "post it": "post-it",
    "post-it": "post-it",
    "scotch tape": "tape",
    "highlighter": "marker",
    "writing tool": "pen",
    "dry erase marker": "whiteboard marker",
    "binder clip": "clip",
    # Add more as needed
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
st.title("\ud83d\udcc2 Office Supply Search")

query = st.text_input("What supply are you looking for? (e.g. pen, sticky note, binder)")

if query:
    df = load_data()
    # Preprocess query using synonym map
    normalized_query = synonym_map.get(query.lower(), query.lower())
    results = fuzzy_search(df, normalized_query)

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
