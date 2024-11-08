import pandas as pd
import requests
from bs4 import BeautifulSoup

def scrape_data_from_url(url):
    # Send a GET request to the URL
    response = requests.get(url)
    
    # Check if the request was successful
    if response.status_code != 200:
        print(f"Failed to retrieve data: {response.status_code}")
        return pd.DataFrame()  # Return an empty DataFrame if not successful
    
    # Parse the HTML content
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # Initialize a list to hold all entries
    heat_data = []

    # Find the main results table
    results_table = soup.find('div', id='divRRPublish')

    if not results_table:
        print("Results table not found!")
        return pd.DataFrame()  # Return an empty DataFrame if not found

    # Extract the data body
    data_body = results_table.find('tbody', id='tb_1_1Data')  # This holds participant data

    if not data_body:
        print("Data body not found!")
        return pd.DataFrame()

    # Extract all participant rows
    rows = data_body.find_all('tr', class_='Hover LastRecordLine')

    for row in rows:
        columns = row.find_all('td')
        if len(columns) >= 8:  # Check if there are enough columns to extract data
            platz = columns[1].get_text(strip=True)  # Position
            name = columns[2].get_text(strip=True)  # Name
            jahrgang = columns[3].get_text(strip=True)  # Year of birth
            nation_img = columns[4].find('img')  # Nation flag image
            nation = nation_img['src'].split('/')[-1].split('.')[0] if nation_img else "Unknown"
            verein = columns[6].get_text(strip=True)  # Club
            zeit = columns[7].get_text(strip=True) if len(columns) > 7 else "No Time"  # Time
            
            # Append the extracted data to heat_data
            heat_data.append({
                'Platz': platz,
                'Name': name,
                'Jahrgang': jahrgang,
                'Nation': nation,
                'Verein': verein,
                'Zeit': zeit,
            })

    # Convert to DataFrame
    df_results = pd.DataFrame(heat_data)
    
    return df_results

if __name__ == "__main__":
    # Provide the URL of the results page you want to scrape
    url = "https://my.raceresult.com/269418/results#240_541D19"  # Change this to your actual URL
    df_results = scrape_data_from_url(url)
    
    if not df_results.empty:
        print("Extracted Results:\n", df_results.head(), "\n")
    else:
        print("No results found.")
