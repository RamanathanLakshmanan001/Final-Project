import pandas as pd
import time
import random
import requests

def search_company_website_with_api(company_name):
    """
    Searches for the official website of a company using SerpAPI.
    """
    API_KEY = "44b9858165582ddc4e7c18e7e8df619999a51e066a250a62c8583b85a95fee62"  # Replace with your SerpAPI key
    query = f"{company_name} official website"
    url = "https://serpapi.com/search"
    
    params = {
        "q": query,
        "hl": "en",
        "gl": "us",
        "api_key": API_KEY
    }
    
    try:
        response = requests.get(url, params=params)
        response.raise_for_status()
        results = response.json()
        
        # Extract the first organic result
        for result in results.get("organic_results", []):
            link = result.get("link")
            if link:
                return link
    except Exception as e:
        print(f"Error searching for {company_name}: {str(e)}")
        return None
    
    return None

def extract_urls_safely(input_file, output_file, max_retries=3):
    """
    Extracts official URLs for companies listed in an Excel file and saves them to a CSV file.
    """
    try:
        # Read the input Excel file
        df = pd.read_excel(input_file)
        
        if 'Company Name' not in df.columns:
            raise ValueError("Input Excel must have a 'Company Name' column")
            
        # Prepare a DataFrame to store results
        results = pd.DataFrame(columns=['Company Name', 'Official URL'])
        
        for index, row in df.iterrows():
            company = row['Company Name']
            print(f"Processing: {company}")
            
            url = None
            for attempt in range(max_retries):
                try:
                    url = search_company_website_with_api(company)
                    if url:
                        break
                except Exception as e:
                    if attempt == max_retries - 1:
                        print(f"Failed after {max_retries} attempts for {company}")
                    time.sleep(random.uniform(5, 15))  # Random delay between attempts
            
            results.loc[index] = [company, url]
            
            # Random delay between companies (5-15 seconds)
            time.sleep(random.uniform(5, 15))
        
        # Save results to a CSV file
        results.to_csv(output_file, index=False)
        print(f"Results saved to {output_file}")
        
    except Exception as e:
        print(f"Error processing files: {e}")

# Example usage
if __name__ == "__main__":
    extract_urls_safely("company_names.xlsx", "company_urls.csv")