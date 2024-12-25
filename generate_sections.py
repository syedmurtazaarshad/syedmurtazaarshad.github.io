import pandas as pd
import os

def generate_publications_section():
    # Load publications.xlsx
    if not os.path.exists('publications.xlsx'):
        print("Error: publications.xlsx not found!")
        return ""
    
    publications = pd.read_excel('publications.xlsx')
    
    # Define the desired order for publication types
    section_order = ["Journal Articles", "Preprints", "Abstracts"]
    output = "<section class=\"publications\">\n<h2>Publications</h2>\n"
    
    for pub_type in section_order:
        group = publications[publications['Type'] == pub_type]
        if not group.empty:
            # Sort by date, latest first
            group = group.sort_values(by='Date', ascending=False)
            
            output += f"<h3>{pub_type}</h3>\n<ul class=\"publication-list\">\n"
            for _, row in group.iterrows():
                published_in = f"<em>{row['Published In']}</em>" if 'Published In' in row and not pd.isna(row['Published In']) else ""
                authors = row['Authors'].replace("Syed M. Arshad", "<strong>Syed M. Arshad</strong>")
                year = pd.to_datetime(row['Date']).year if not pd.isna(row['Date']) else ""
                output += (
                    f"<li>"
                    f"<div class=\"publication-title\"><a href=\"{row['Link']}\" target=\"_blank\">{row['Title']}</a>"
                    f"<span class=\"publication-year\">{year}</span></div>"
                    f"<div class=\"publication-authors\">{authors}</div>"
                    f"<div class=\"publication-journal\">{published_in}</div>"
                    f"</li>\n"
                )
            output += "</ul>\n"
    output += "</section>\n"
    return output

def generate_talks_section():
    # Load talks.xlsx
    if not os.path.exists('talks.xlsx'):
        print("Error: talks.xlsx not found!")
        return ""
    
    talks = pd.read_excel('talks.xlsx')
    
    # Sort by date, latest first
    talks = talks.sort_values(by='Date', ascending=False)
    
    output = "<section class=\"talks\">\n<h2>Talks & Presentations</h2>\n<ul class=\"talk-list\">\n"
    for _, row in talks.iterrows():
        year = pd.to_datetime(row['Date']).year if not pd.isna(row['Date']) else ""
        details = f"<em>{row['Details']}</em>" if 'Details' in row and not pd.isna(row['Details']) else ""
        
        if pd.notna(row['Link']):
            output += (
                f"<li>"
                f"<div class=\"talk-title\"><a href=\"{row['Link']}\" target=\"_blank\">{row['Title']}</a>"
                f"<span class=\"talk-year\">{year}</span></div>"
                f"<div class=\"talk-details\">{details}</div>"
                f"</li>\n"
            )
        else:
            output += (
                f"<li>"
                f"<div class=\"talk-title\">{row['Title']}"
                f"<span class=\"talk-year\">{year}</span></div>"
                f"<div class=\"talk-details\">{details}</div>"
                f"</li>\n"
            )
    output += "</ul>\n</section>\n"
    return output

if __name__ == "__main__":
    publications_section = generate_publications_section()
    talks_section = generate_talks_section()
    
    output_path = os.path.join(os.getcwd(), 'generated_sections.md')
    with open(output_path, 'w', encoding='utf-8') as f:
        if publications_section:
            f.write(publications_section)
        if talks_section:
            f.write(talks_section)
    
    print(f"Web content generated successfully at {output_path}!")
