import requests
from bs4 import BeautifulSoup
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import time
import os
import random
from urllib.parse import quote_plus

# Create default output directory
DEFAULT_OUTPUT_DIR = os.path.join(os.getcwd(), "Output")
if not os.path.exists(DEFAULT_OUTPUT_DIR):
    os.makedirs(DEFAULT_OUTPUT_DIR)

# List of User-Agents to rotate through
USER_AGENTS = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Edge/91.0.864.59 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.1 Safari/605.1.15',
    'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:91.0) Gecko/20100101 Firefox/91.0'
]

def scrape_scholar_articles(query, num_pages):
    articles = []
    page = 0
    
    # Add user agent to mimic a browser
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    while page < num_pages:
        try:
            url = f"https://scholar.google.com/scholar?start={page*10}&q={query}&hl=en&as_sdt=0,5"
            print(f"Fetching page {page + 1}...")
            
            response = requests.get(url, headers=headers)
            if response.status_code != 200:
                messagebox.showerror("Error", f"Failed to fetch results. Status code: {response.status_code}")
                break
                
            soup = BeautifulSoup(response.text, "html.parser")
            results = soup.find_all("div", class_="gs_ri")
            
            if not results:
                print("No results found on this page.")
                messagebox.showwarning("Warning", "No results found. Google Scholar might be blocking automated access.")
                break
            
            for result in results:
                try:
                    title_elem = result.find("h3", class_="gs_rt")
                    title = title_elem.text if title_elem else "No title found"
                    
                    authors_elem = result.find("div", class_="gs_a")
                    authors = authors_elem.text if authors_elem else "No authors found"
                    
                    link_elem = title_elem.find("a") if title_elem else None
                    link = link_elem["href"] if link_elem and "href" in link_elem.attrs else "No link found"
                    
                    articles.append({"Title": title, "Authors": authors, "Link": link})
                    print(f"Found article: {title[:50]}...")
                except Exception as e:
                    print(f"Error parsing result: {str(e)}")
            
            # Add a delay between requests to avoid being blocked
            time.sleep(2)
            page += 1
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            break
    
    return articles

def save_to_excel(articles, filename):
    if not articles:
        messagebox.showwarning("Warning", "No articles to save!")
        return
        
    try:
        # Check if file exists and generate new filename if it does
        base_filename = os.path.splitext(filename)[0]
        extension = os.path.splitext(filename)[1]
        counter = 1
        new_filename = filename
        
        while os.path.exists(new_filename):
            new_filename = f"{base_filename}_{counter}{extension}"
            counter += 1
            
        df = pd.DataFrame(articles)
        df.to_excel(new_filename, index=False)
        messagebox.showinfo("Success", f"Successfully saved {len(articles)} articles to {new_filename}")
        return new_filename
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save file: {str(e)}")
        return None

def browse_folder():
    folder_path = filedialog.askdirectory(initialdir=DEFAULT_OUTPUT_DIR)
    entry_folder.delete(0, tk.END)
    entry_folder.insert(tk.END, folder_path)

def scrape_articles():
    query = entry_query.get()
    try:
        num_pages = int(entry_pages.get())
    except ValueError:
        messagebox.showerror("Error", "Please enter a valid number of pages")
        return

    articles = scrape_scholar_articles(query, num_pages)

    folder_path = entry_folder.get()
    if not folder_path:  # If no folder is selected, use default
        folder_path = DEFAULT_OUTPUT_DIR
    
    filename = os.path.join(folder_path, "scholar_articles.xlsx")
    saved_filename = save_to_excel(articles, filename)
    if saved_filename:
        label_status.config(text=f"Extraction complete. Data saved to {os.path.basename(saved_filename)}")
    else:
        label_status.config(text="Failed to save data. Please check the error message.")

# Create the main window
window = tk.Tk()
window.title("Google Scholar Scraper")
window.geometry("400x250")

# Create input fields and labels
label_query = tk.Label(window, text="Article Title or Keyword:")
label_query.pack()
entry_query = tk.Entry(window, width=40)
entry_query.pack()

label_pages = tk.Label(window, text="Number of Pages:")
label_pages.pack()
entry_pages = tk.Entry(window, width=40)
entry_pages.pack()

label_folder = tk.Label(window, text="Output Folder (optional):")
label_folder.pack()
entry_folder = tk.Entry(window, width=40)
entry_folder.insert(0, DEFAULT_OUTPUT_DIR)  # Set default output directory in the entry field
entry_folder.pack()

# Create browse button
button_browse = tk.Button(window, text="Browse", command=browse_folder)
button_browse.pack()

# Create extract button
button_extract = tk.Button(window, text="Extract Data", command=scrape_articles)
button_extract.pack()

# Create status label
label_status = tk.Label(window, text="")
label_status.pack()

# Run the main window loop
window.mainloop()

class ScholarScraper:
    def __init__(self):
        # ... existing init code ...
        self.stats['blocked_requests'] = 0  # Add new statistic
        # ... rest of init ...

    # ... existing methods until scrape_scholar ...

    def get_random_wait_time(self):
        """Get a random wait time between requests"""
        # Base delay between 5 and 10 seconds
        base_delay = random.uniform(5, 10)
        
        # 20% chance of longer delay (15-30 seconds)
        if random.random() < 0.2:
            base_delay = random.uniform(15, 30)
            
        # 5% chance of very long delay (45-90 seconds)
        if random.random() < 0.05:
            base_delay = random.uniform(45, 90)
            
        return base_delay

    def make_request(self, url, retry_count=0, max_retries=3):
        """Make a request with random delays and retry logic"""
        if retry_count >= max_retries:
            return None

        # Random User-Agent
        headers = {
            'User-Agent': random.choice(USER_AGENTS),
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
            'TE': 'Trailers',
            'DNT': '1'  # Do Not Track header
        }

        # Random wait before request
        wait_time = self.get_random_wait_time()
        self.progress_var.set(f"Waiting {wait_time:.1f} seconds before next request...")
        self.window.update()
        time.sleep(wait_time)

        try:
            response = requests.get(url, headers=headers, timeout=30)
            
            if response.status_code == 200:
                return response
            elif response.status_code == 429 or response.status_code == 403:
                self.stats['blocked_requests'] += 1
                # Exponential backoff with randomization
                wait_time = (2 ** retry_count) * random.uniform(20, 30)
                self.progress_var.set(f"Access blocked. Waiting {wait_time:.1f} seconds before retry...")
                self.window.update()
                time.sleep(wait_time)
                return self.make_request(url, retry_count + 1, max_retries)
            else:
                return None

        except Exception as e:
            print(f"Request error: {str(e)}")
            return None

    def scrape_scholar(self):
        articles = []
        seen_titles = set()
        
        # Get all non-empty keywords
        keywords = ' '.join(entry.get() for entry in self.keywords if entry.get().strip())
        
        if not keywords:
            messagebox.showerror("Error", "Please enter at least one keyword")
            return []
        
        try:
            num_pages = int(self.num_pages.get())
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid number of pages")
            return []
        
        self.progress_bar['maximum'] = num_pages
        page = 0
        
        while page < num_pages:
            # Encode keywords properly
            encoded_keywords = quote_plus(keywords)
            url = f"https://scholar.google.com/scholar?start={page*10}&q={encoded_keywords}&hl=en&as_sdt=0,5"
            
            self.progress_var.set(f"Fetching page {page + 1}...")
            self.progress_bar['value'] = page + 1
            self.window.update()
            
            response = self.make_request(url)
            if response is None:
                messagebox.showerror("Error", "Failed to fetch results after multiple retries")
                break
            
            soup = BeautifulSoup(response.text, "html.parser")
            results = soup.find_all("div", class_="gs_ri")
            
            if not results:
                print("No results found on this page.")
                if page == 0:
                    messagebox.showwarning("Warning", "No results found for the given keywords.")
                else:
                    messagebox.showinfo("Info", "No more results available.")
                break
            
            # ... rest of the scraping logic ...
            
            # Add some randomization to the page progression
            if random.random() < 0.1:  # 10% chance of longer delay between pages
                extra_delay = random.uniform(10, 20)
                self.progress_var.set(f"Adding extra delay of {extra_delay:.1f} seconds...")
                self.window.update()
                time.sleep(extra_delay)
            
            page += 1
        
        if self.stats['blocked_requests'] > 0:
            messagebox.showinfo("Access Blocks", 
                              f"Encountered {self.stats['blocked_requests']} access blocks during scraping.\n"
                              "Consider reducing the number of pages or increasing delays between requests.")
        
        return articles

    def update_stats(self):
        stats_text = f"""
Articles Found: {self.stats['total_found']}
Excluded - No Full Text: {self.stats['excluded_no_fulltext']}
Excluded - Non-English: {self.stats['excluded_non_english']}
Excluded - Before 2018: {self.stats['excluded_old_date']}
Excluded - Duplicates: {self.stats['excluded_duplicate']}
Access Blocks: {self.stats['blocked_requests']}
Included Articles: {self.stats['included_final']}
        """
        self.stats_text.delete(1.0, tk.END)
        self.stats_text.insert(tk.END, stats_text)

    # ... rest of the class methods ...
