import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
from googlesearch import search
import time
import os
from colorama import init, Fore, Style

def load_progress():
    if os.path.exists('progress.txt'):
        # Try different character encodings
        encodings = ['utf-8', 'utf-8-sig', 'latin1', 'cp1252']
        for encoding in encodings:
            try:
                with open('progress.txt', 'r', encoding=encoding) as f:
                    content = f.read().strip()
                    # Return None if file is empty or contains only whitespace
                    return content if content else None
            except UnicodeDecodeError:
                continue
            except Exception as e:
                print(f"{Fore.YELLOW}Error reading progress file: {str(e)}{Style.RESET_ALL}")
                continue
        
        # If no encoding worked
        print(f"{Fore.YELLOW}Could not read progress file. Starting from beginning.{Style.RESET_ALL}")
        return None
    return None

def save_progress(company):
    try:
        with open('progress.txt', 'w', encoding='utf-8') as f:
            f.write(company)
    except Exception as e:
        print(f"{Fore.YELLOW}Error saving progress: {str(e)}{Style.RESET_ALL}")

def update_excel(company, email, status, target_country):
    filename = 'email_results.xlsx'
    
    if os.path.exists(filename):
        df = pd.read_excel(filename)
    else:
        df = pd.DataFrame(columns=['Company', 'Email', 'Status', 'Target Country'])
    
    new_row = pd.DataFrame({
        'Company': [company], 
        'Email': [email], 
        'Status': [status],
        'Target Country': [target_country]
    })
    df = pd.concat([df, new_row], ignore_index=True)
    df.to_excel(filename, index=False)

def find_email(text):
    # More comprehensive email pattern
    email_patterns = [
        r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}',  # Standard email
        r'(?:mailto:)?([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,})',  # mailto: links
        r'[A-Za-z0-9._%+-]+\s*[\[\(]at\[\)\]][A-Za-z0-9.-]+\.[A-Z|a-z]{2,}',  # at/@ variations
    ]
    
    found_emails = []
    for pattern in email_patterns:
        emails = re.findall(pattern, text, re.IGNORECASE)
        found_emails.extend(emails)
    
    # Email validation and cleaning
    valid_emails = []
    for email in found_emails:
        email = email.strip().lower()
        if '@' in email and '.' in email.split('@')[1]:
            # Filter general/spam emails
            if not any(spam in email for spam in ['example.com', 'test.com']):
                valid_emails.append(email)
    
    return valid_emails[0] if valid_emails else None

def extract_emails_from_html(soup):
    emails = set()
    
    # Collect emails from links
    for link in soup.find_all('a'):
        href = link.get('href', '')
        if 'mailto:' in href:
            email = href.replace('mailto:', '').strip()
            emails.add(email)
    
    # Check elements with specific classes and IDs
    contact_elements = soup.find_all(class_=lambda x: x and ('contact' in x.lower() or 
                                                           'email' in x.lower() or 
                                                           'impressum' in x.lower()))
    for element in contact_elements:
        text = element.get_text()
        found = find_email(text)
        if found:
            emails.add(found)
    
    return list(emails)

def clean_company_name(company):
    """Cleans special characters from company name, preserves important ones"""
    # Allowed characters: letters, numbers, +, -, &, space
    cleaned = re.sub(r'[^a-zA-Z0-9\s\+\-\&]', ' ', company)
    # Reduce multiple spaces to single space
    cleaned = re.sub(r'\s+', ' ', cleaned)
    return cleaned.strip()

def scrape_company_email(company):
    try:
        # Clean company name
        clean_company = clean_company_name(company)
        print(f"Cleaned company name: {clean_company}")
        
        # First phase - specific searches
        search_queries = [
            f"{clean_company} contact email",
            f"{clean_company} kontakt email",
            f"{clean_company} impressum",
            f"{clean_company} about us",
            f"site:{clean_company.lower().replace(' ', '')}.com contact",
            f"site:{clean_company.lower().replace(' ', '')}.de contact"
        ]
        
        # First attempt to find email
        for search_query in search_queries:
            email = try_search_query(search_query)
            if email:
                return email
        
        # Second phase - general email search
        print("No email found in first phase, performing general search...")
        general_queries = [
            f"{clean_company} email",
            f"{clean_company} mail",
            f"{clean_company} info@",
            f"{clean_company} contact@"
        ]
        
        for search_query in general_queries:
            # Check more results
            search_results = search(search_query, 
                                  num=5,  # Increased number of results
                                  stop=5, 
                                  pause=2.0,
                                  lang='en')
            
            for url in search_results:
                if any(skip_domain in url.lower() for skip_domain in [
                    'linkedin.com', 'facebook.com', 'twitter.com', 
                    'instagram.com', 'youtube.com'
                ]):
                    continue  # Skip social media sites
                    
                try:
                    headers = {
                        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                    }
                    response = requests.get(url, timeout=10, headers=headers)
                    soup = BeautifulSoup(response.text, 'html.parser')
                    
                    # Extract email from HTML
                    html_emails = extract_emails_from_html(soup)
                    if html_emails:
                        # Try to find email related to company name
                        company_name_parts = clean_company.lower().split()
                        for email in html_emails:
                            if any(part in email.lower() for part in company_name_parts):
                                return email
                        # If no email matches company name, return first found email
                        return html_emails[0]
                    
                    # Text-based search
                    page_text = soup.get_text()
                    email = find_email(page_text)
                    if email:
                        return email
                        
                except Exception as e:
                    print(f"URL scanning error: {str(e)}")
                    continue
        
        return None
        
    except Exception as e:
        print(f"Search error: {str(e)}")
        return None

def try_search_query(search_query):
    """Search for email with a single query"""
    print(f"Trying query: {search_query}")
    try:
        # Error handling and retry logic
        max_retries = 3
        for attempt in range(max_retries):
            try:
                search_results = search(search_query, num=3, stop=3, pause=2.0, lang='en')
                # Convert generator to list to catch errors immediately
                urls = list(search_results)
                break
            except Exception as e:
                if "429" in str(e):  # Rate limit error
                    print(f"{Fore.YELLOW}⚠️ Google query limit exceeded. Waiting 60 seconds...{Style.RESET_ALL}")
                    time.sleep(60)
                    if attempt == max_retries - 1:
                        raise
                else:
                    raise
        
        for url in urls:
            print(f"Scanning URL: {url}")
            try:
                headers = {
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                }
                response = requests.get(url, timeout=10, headers=headers)
                soup = BeautifulSoup(response.text, 'html.parser')
                
                html_emails = extract_emails_from_html(soup)
                if html_emails:
                    return html_emails[0]
                
                page_text = soup.get_text()
                email = find_email(page_text)
                if email:
                    return email
                    
            except Exception as e:
                print(f"URL scanning error: {str(e)}")
                continue
                
        return None
    except Exception as e:
        print(f"{Fore.RED}Query error: {str(e)}{Style.RESET_ALL}")
        return None

def get_source_file():
    print(f"{Fore.CYAN}Files in current directory:{Style.RESET_ALL}")
    
    # List all files in directory
    files = [f for f in os.listdir('.') if os.path.isfile(f)]
    
    # Show numbered files
    for i, file in enumerate(files, 1):
        print(f"{i}. {file}")
    
    while True:
        try:
            choice = input(f"\n{Fore.CYAN}Which file contains the company list? (Enter number): {Style.RESET_ALL}")
            file_index = int(choice) - 1
            
            if 0 <= file_index < len(files):
                selected_file = files[file_index]
                print(f"\n{Fore.GREEN}Selected file: {selected_file}{Style.RESET_ALL}")
                return selected_file
            else:
                print(f"{Fore.RED}Invalid number! Please enter a number between 1 and {len(files)}.{Style.RESET_ALL}")
        except ValueError:
            print(f"{Fore.RED}Please enter a valid number!{Style.RESET_ALL}")

def main():
    # Colorama'yı başlat
    init()
    
    try:
        # Kaynak dosyayı seç
        source_file = get_source_file()
        
        # Dosya adından ülke adını çıkar (örn: "almanya.txt" -> "Almanya")
        target_country = os.path.splitext(os.path.basename(source_file))[0].capitalize()
        print(f"\n{Fore.GREEN}Target Country: {target_country}{Style.RESET_ALL}")
        
        # Excel dosyasını kontrol et
        existing_companies = set()
        if os.path.exists('email_results.xlsx'):
            df = pd.read_excel('email_results.xlsx')
            existing_companies = set(df['Company'].tolist())
        
        # Seçilen dosyadan şirket listesini oku
        try:
            with open(source_file, 'r', encoding='utf-8') as f:
                companies = f.read().splitlines()
        except UnicodeDecodeError:
            # UTF-8 ile açılamazsa diğer kodlamaları dene
            encodings = ['utf-8-sig', 'latin1', 'cp1252']
            for encoding in encodings:
                try:
                    with open(source_file, 'r', encoding=encoding) as f:
                        companies = f.read().splitlines()
                    break
                except UnicodeDecodeError:
                    continue
        
        # Son kaldığımız yeri kontrol et
        last_company = load_progress()
        if last_company:
            try:
                start_index = companies.index(last_company) + 1
            except ValueError:
                print(f"{Fore.YELLOW}Warning: Last processed company not found in list. Starting from beginning.{Style.RESET_ALL}")
                start_index = 0
        else:
            start_index = 0
        
        total_companies = len(companies[start_index:])
        
        skipped_companies = []  # List to track skipped companies
        
        for index, company in enumerate(companies[start_index:], 1):
            if not company or company.isspace():  # Check for empty or whitespace-only lines
                print(f"{Fore.YELLOW}⚠️ Skipped empty line (line {index}){Style.RESET_ALL}")
                skipped_companies.append((index, company, "Empty line"))
                continue
            
            try:
                print("\n" + "="*50)
                print(f"{Fore.CYAN}📋 Company {index}/{total_companies} - {company}{Style.RESET_ALL}")
                print("="*50)
                
                # Şirket daha önce aranmış mı kontrol et
                if company in existing_companies:
                    print(f"{Fore.YELLOW}⚠️ This company has already been searched. Skipping...{Style.RESET_ALL}")
                    continue
                
                # Şirket adı geçerlilik kontrolü
                if len(company) < 2:  # Check for very short company names
                    print(f"{Fore.YELLOW}⚠️ Invalid company name: '{company}'{Style.RESET_ALL}")
                    status = "Invalid company name"
                    skipped_companies.append((index, company, status))
                    update_excel(company, "-", status, target_country)
                    save_progress(company)
                    continue
                
                try:
                    email = scrape_company_email(company)
                except Exception as scrape_error:
                    error_msg = str(scrape_error)
                    print(f"{Fore.RED}Email search error: {error_msg}{Style.RESET_ALL}")
                    status = f"Email search error: {error_msg}"
                    skipped_companies.append((index, company, status))
                    email = "-"
                    update_excel(company, email, status, target_country)
                    save_progress(company)
                    time.sleep(2)
                    continue
                
                if email:
                    print(f"{Fore.GREEN}✅ Email found: {email}{Style.RESET_ALL}")
                    status = "Email found"
                else:
                    print(f"{Fore.RED}❌ Email not found{Style.RESET_ALL}")
                    status = "Email not found"
                    email = "-"
                
                try:
                    update_excel(company, email, status, target_country)
                    save_progress(company)
                except Exception as save_error:
                    print(f"{Fore.RED}Error saving: {str(save_error)}{Style.RESET_ALL}")
                
            except Exception as e:
                error_msg = str(e)
                print(f"{Fore.RED}Unexpected error: {error_msg}{Style.RESET_ALL}")
                print(f"{Fore.RED}Error details:{Style.RESET_ALL}")
                import traceback
                print(traceback.format_exc())
                
                status = f"Critical error: {error_msg}"
                skipped_companies.append((index, company, status))
                
                try:
                    update_excel(company, "-", status, target_country)
                    save_progress(company)
                except:
                    print(f"{Fore.RED}Error saving failed!{Style.RESET_ALL}")
            
            finally:
                time.sleep(2)
        
        # Program sonunda atlanan şirketleri raporla
        if skipped_companies:
            print("\n" + "="*50)
            print(f"{Fore.YELLOW}⚠️ Skipped or Error Companies:{Style.RESET_ALL}")
            print("="*50)
            for idx, comp, reason in skipped_companies:
                print(f"{Fore.YELLOW}Index: {idx}, Company: {comp}, Reason: {reason}{Style.RESET_ALL}")
            
            # Atlanan şirketleri ayrı bir dosyaya kaydet
            try:
                with open('skipped_companies.txt', 'w', encoding='utf-8') as f:
                    f.write("Index\tCompany\tReason\n")
                    for idx, comp, reason in skipped_companies:
                        f.write(f"{idx}\t{comp}\t{reason}\n")
                print(f"\n{Fore.YELLOW}Skipped companies saved to 'skipped_companies.txt' file.{Style.RESET_ALL}")
            except Exception as save_error:
                print(f"{Fore.RED}Error saving skipped companies: {str(save_error)}{Style.RESET_ALL}")

        print(f"\n{Fore.GREEN}✨ All companies processed!{Style.RESET_ALL}")
        print(f"Total skipped/error companies: {len(skipped_companies)}")

    except Exception as e:
        print(f"\n{Fore.RED}Critical program error: {str(e)}{Style.RESET_ALL}")
        print(f"{Fore.RED}Error details:{Style.RESET_ALL}")
        import traceback
        print(traceback.format_exc())

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n{Fore.YELLOW}Program interrupted by user.{Style.RESET_ALL}")
    except Exception as e:
        print(f"\n{Fore.RED}Critical program error: {str(e)}{Style.RESET_ALL}")
        print(f"{Fore.RED}Error details:{Style.RESET_ALL}")
        import traceback
        print(traceback.format_exc()) 