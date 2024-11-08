import requests
from bs4 import BeautifulSoup
import pandas as pd

def extract_headers(block_header):
    headers = []
    for header_div in block_header.find_all("div", recursive=False):
        first_line_text = header_div.find_all("div")[0].get_text(strip=True)
        second_line = header_div.find("div", class_="secondline")
        
        if second_line:
            second_line_text = second_line.get_text(strip=True)
            headers.append((first_line_text, second_line_text))  # Tuple with first and second line
        else:
            headers.append((first_line_text,))  # Tuple with only first line
    
    return headers



def scrape_dlv_data(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    
    dataframes = {}
    heat_blocks = soup.find_all(class_=lambda x: x and (x.startswith("runblock heatblock") or x.startswith("runblock roundblock") or x.startswith("startlistblock")))
    
    for heat in heat_blocks:
        blockname = heat.find(class_="blockname")
        leftname = blockname.find(class_="leftname") if blockname else None
        heat_name = leftname.get_text(strip=True) if leftname else "Unknown Heat"
        
        result_blocks = heat.find_all(class_="resultblock")
        
        for block in result_blocks:
            block_table = block.find(class_="blocktable")
            block_header = block_table.find(class_="blockheader")
            headers = extract_headers(block_header)
            
            entries = block_table.find_all("div", recursive=False)[1:]  # Skipping the blockheader
            heat_data = []
            
            for entry in entries:
                entry_data = {}
                columns = entry.find_all("div", recursive=False)

                for i, header_tuple in enumerate(headers):
                    column_data = columns[i].find_all("div")
                    entry_data[header_tuple[0]] = column_data[0].get_text(" ", strip=True) if column_data else ""
                    
                    if len(header_tuple) == 2 and len(column_data) > 1:
                        entry_data[header_tuple[1]] = column_data[1].get_text(" ", strip=True)

                heat_data.append(entry_data)
                
            dataframes[heat_name] = pd.DataFrame(heat_data)
    return dataframes

if __name__ == "__main__":
    url = "https://ergebnisse.leichtathletik.de/Competitions/CurrentList/617972/12005"
    dataframes = scrape_dlv_data(url)
    for heat_name, df in dataframes.items():
        print(f"Heat: {heat_name}\n", df.head(), "\n")
