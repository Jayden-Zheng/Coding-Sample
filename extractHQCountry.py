import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
from bs4 import BeautifulSoup

# Define the function to fetch the text from the URL
def fetch_text(url, phrase):
    response = requests.get(url)
    content = response.text
    soup = BeautifulSoup(content, 'html.parser')

    # Get the text content without HTML tags
    text_content = soup.get_text()

    # Split content into sentences
    sentences = text_content.split('. ')

    # Search for the phrase within each sentence
    for sentence in sentences:
        if phrase.lower() in sentence.lower():
            # Get the text until the first occurrence of a colon, "Hours," or "Phone"
            stop_words = ['Address', 'Hours', 'Phone', 'Founded', 'Number', ':', 'in']
            split_result = sentence.split(phrase, 1)
            if len(split_result) > 1:
                result = split_result[1]
                for word in stop_words:
                    if word.lower() in result.lower():
                        result = result.split(word, 1)[0]
                        break
                return result.strip()

    # If the phrase is not found, search for "Headquarters:" and extract the first 4 words after it
    if phrase.lower() == 'address:':
        return fetch_text(url, 'Headquarters:').split()[:4]

    # Return an empty string if the phrase and "Headquarters:" are not found
    return ""

# Set up credentials and authorize the API client
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']      #hard-coded
credentials = ServiceAccountCredentials.from_json_keyfile_name('C:/Users/ZHENGF7/Downloads/global-org-location-analysis-ce51d92cd9ef.json', scope)    #hard-coded
client = gspread.authorize(credentials)

# Open the Google Sheet
sheet = client.open('SHEET #1').worksheet('LOCATION MAPPING SHEET')  # Replace with your sheet name

# Get the URLs from column B
urls = sheet.col_values(2)[1:]  # Skip the header row

# Apply the function and update the cells in column C
# If starting from the beginning, just put urls as the argument. And then make "i+2" for the i below.
for i, url in enumerate(urls):
#for i, url in enumerate(urls[1137:], start=1138):
    phrase = 'Address:'  # Modify the desired phrase as needed
    result = fetch_text(url, phrase)
    print(f"Updating cell {i+2}, 3 with result: {result}")
    #print(f"Updating cell {i+1}, 3 with result: {result}")

    # Handle empty or invalid result values
    if result:
        if isinstance(result, list):
            result = ' '.join(result)  # Convert list to string
        sheet.update_cell(i+2, 3, result)  # Update column C, starting from row 2 (!!!!!!! Double check if it aligns with the "Updating cells index above" and if it produces output in the correct row!!!!!!!!)
    else:
        sheet.update_cell(i+2, 3, "")  # Update with an empty string if result is empty

print("Cells updated successfully.")