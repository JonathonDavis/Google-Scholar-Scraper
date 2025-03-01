import requests
from bs4 import BeautifulSoup
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import time
import os

# Create default output directory
DEFAULT_OUTPUT_DIR = os.path.join(os.getcwd(), "Output")
if not os.path.exists(DEFAULT_OUTPUT_DIR):
    os.makedirs(DEFAULT_OUTPUT_DIR)

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
