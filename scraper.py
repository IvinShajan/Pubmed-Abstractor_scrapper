from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
from datetime import datetime
import os
import sys

class Config:
    INPUT_FILE = "input.xlsx"
    OUTPUT_FILE = "output.xlsx"
    USER_NAME = "IvinShajan"
    MAX_PAPERS_PER_KINASE = 10
    WAIT_TIME = 20
    PAGE_LOAD_DELAY = 3
    MAX_RETRIES = 3

class PubMedKinaseScraper:
    def __init__(self, input_file_path, output_file_path, user_name=Config.USER_NAME, current_datetime=None):
        self.input_file_path = input_file_path
        self.output_file_path = output_file_path
        self.user_name = user_name
        self.current_datetime = current_datetime if current_datetime else datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')
        self.driver = None
        self.results = []

    def setup_driver(self):
        """Setup Chrome WebDriver with additional options"""
        options = Options()
        # Comment out headless mode for debugging
        # options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--start-maximized')
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-infobars')
        options.add_argument('--disable-notifications')
        
        try:
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.driver.implicitly_wait(Config.WAIT_TIME)
            return True
        except Exception as e:
            print(f"Failed to setup WebDriver: {e}")
            return False

    def load_kinases(self, kinase_column=None):
        """Load kinase names from Excel file"""
        try:
            print(f"Attempting to load kinases from: {self.input_file_path}")
            df = pd.read_excel(self.input_file_path)
            print(f"Available columns: {df.columns.tolist()}")
            
            if kinase_column and kinase_column in df.columns:
                kinase_column = kinase_column
            else:
                kinase_column = None
                for col in df.columns:
                    if 'kinase' in col.lower():
                        kinase_column = col
                        break

            if kinase_column is None:
                raise ValueError("No column with 'kinase' in its name found. Please specify the column name.")

            kinases = df[kinase_column].dropna().tolist()
            print(f"Successfully loaded {len(kinases)} kinases")
            return kinases
        except Exception as e:
            print(f"Error loading kinases from Excel file: {e}")
            return []

    def retry_operation(self, operation, max_retries=Config.MAX_RETRIES):
        """Retry an operation with exponential backoff"""
        for attempt in range(max_retries):
            try:
                return operation()
            except Exception as e:
                wait_time = (2 ** attempt) + 1
                if attempt < max_retries - 1:
                    print(f"Attempt {attempt + 1} failed. Retrying in {wait_time} seconds...")
                    time.sleep(wait_time)
                else:
                    print(f"All {max_retries} attempts failed. Last error: {e}")
                    return None

    def search_pubmed(self, kinase):
        """Search PubMed for a specific kinase"""
        def _search():
            base_url = "https://pubmed.ncbi.nlm.nih.gov/"
            self.driver.get(base_url)
            time.sleep(Config.PAGE_LOAD_DELAY)

            search_box = WebDriverWait(self.driver, Config.WAIT_TIME).until(
                EC.presence_of_element_located((By.NAME, "term"))
            )
            search_box.clear()
            search_query = f"({kinase}[Title/Abstract]) AND (human[Title/Abstract] OR humans[MeSH Terms])"
            search_box.send_keys(search_query)
            search_box.send_keys(Keys.RETURN)
            time.sleep(Config.PAGE_LOAD_DELAY * 2)
            return True

        return self.retry_operation(_search)

    def get_total_results(self):
        """Get total number of search results"""
        try:
            results_text = WebDriverWait(self.driver, Config.WAIT_TIME).until(
                EC.presence_of_element_located((By.CLASS_NAME, "results-amount"))
            ).text
            return int(''.join(filter(str.isdigit, results_text)))
        except (TimeoutException, NoSuchElementException, ValueError):
            return 0

    def extract_abstract(self):
        """Extract abstract from the current PubMed page"""
        try:
            # Try to expand the abstract if needed
            try:
                expand_button = self.driver.find_element(By.CLASS_NAME, "abstract-expander")
                expand_button.click()
                time.sleep(1)
            except:
                pass

            # Try multiple selectors for abstract content
            abstract_selectors = [
                "abstract-content.selected",
                "abstract-content",
                "abstract",
                "abstract-1"
            ]

            abstract_text = None
            for selector in abstract_selectors:
                try:
                    abstract_element = WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.CLASS_NAME, selector))
                    )
                    abstract_text = abstract_element.text.strip()
                    if abstract_text:
                        break
                except:
                    continue

            if not abstract_text:
                try:
                    abstract_element = self.driver.find_element(
                        By.XPATH, "//div[contains(@class, 'abstract-content')]//p"
                    )
                    abstract_text = abstract_element.text.strip()
                except:
                    pass

            return abstract_text if abstract_text else "Abstract not found"

        except Exception as e:
            print(f"Error extracting abstract: {e}")
            return "Error extracting abstract"

    def extract_title(self):
        """Extract article title"""
        try:
            title = WebDriverWait(self.driver, Config.WAIT_TIME).until(
                EC.presence_of_element_located((By.CLASS_NAME, "heading-title"))
            )
            return title.text.strip()
        except (TimeoutException, NoSuchElementException):
            return "Title not found"

    def get_all_abstracts(self, kinase):
        """Get all abstracts for a kinase"""
        all_papers = []
        total_results = self.get_total_results()

        if total_results == 0:
            print(f"No results found for {kinase}")
            return []

        print(f"Found {total_results} papers for {kinase}")
        current_page = 1
        processed_count = 0

        while processed_count < Config.MAX_PAPERS_PER_KINASE:
            try:
                article_links = WebDriverWait(self.driver, Config.WAIT_TIME).until(
                    EC.presence_of_all_elements_located((By.CLASS_NAME, "docsum-title"))
                )

                for link in article_links:
                    if processed_count >= Config.MAX_PAPERS_PER_KINASE:
                        break

                    try:
                        # Store the current window handle
                        main_window = self.driver.current_window_handle
                        
                        # Open in new tab and switch to it
                        self.driver.execute_script("arguments[0].setAttribute('target', '_blank');", link)
                        link.click()
                        time.sleep(Config.PAGE_LOAD_DELAY)

                        # Switch to the new tab
                        new_window = [window for window in self.driver.window_handles if window != main_window][0]
                        self.driver.switch_to.window(new_window)

                        title = self.extract_title()
                        abstract = self.extract_abstract()

                        if title and abstract:
                            paper_info = {
                                'title': title,
                                'abstract': abstract
                            }
                            all_papers.append(paper_info)
                            processed_count += 1
                            print(f"Extracted abstract {processed_count} of {min(total_results, Config.MAX_PAPERS_PER_KINASE)} for {kinase}")

                        # Close the tab and switch back
                        self.driver.close()
                        self.driver.switch_to.window(main_window)
                        time.sleep(1)

                    except Exception as e:
                        print(f"Error processing article: {e}")
                        # Ensure we're back on the main window
                        if self.driver.current_window_handle != main_window:
                            self.driver.close()
                            self.driver.switch_to.window(main_window)
                        continue

                if processed_count >= Config.MAX_PAPERS_PER_KINASE:
                    break

                # Try to go to next page
                try:
                    next_button = self.driver.find_element(By.CLASS_NAME, "next-page")
                    if 'disabled' in next_button.get_attribute('class'):
                        break
                    next_button.click()
                    time.sleep(Config.PAGE_LOAD_DELAY)
                    current_page += 1
                except:
                    break

            except Exception as e:
                print(f"Error processing page {current_page}: {e}")
                break

        return all_papers

    def format_papers_for_excel(self, papers):
        """Format papers for Excel output"""
        formatted_text = ""
        for i, paper in enumerate(papers, 1):
            formatted_text += f"[Paper {i}]\n"
            formatted_text += f"Title: {paper['title']}\n"
            formatted_text += f"Abstract: {paper['abstract']}\n"
            formatted_text += "-" * 80 + "\n"
        return formatted_text

    def save_to_excel(self, results, output_file=None):
        """Save results to Excel file"""
        if output_file is None:
            output_file = self.output_file_path
            
        df = pd.DataFrame(results)
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Abstracts', index=False)
            metadata = pd.DataFrame({
                'Metadata': ['Extraction Date (UTC)', 'User', 'Total Kinases Processed'],
                'Value': [self.current_datetime, self.user_name, len(results)]
            })
            metadata.to_excel(writer, sheet_name='Metadata', index=False)

    def save_partial_results(self):
        """Save current results to avoid losing progress"""
        if self.results:
            try:
                temp_output = self.output_file_path.replace('.xlsx', '_partial.xlsx')
                self.save_to_excel(self.results, temp_output)
                print(f"Partial results saved to: {temp_output}")
            except Exception as e:
                print(f"Failed to save partial results: {e}")

    def process_kinases(self, kinase_column=None):
        """Main processing function"""
        if not self.setup_driver():
            print("Failed to initialize WebDriver. Exiting...")
            return

        try:
            kinases = self.load_kinases(kinase_column)
            if not kinases:
                print("No kinases found in input file. Exiting...")
                return

            print(f"Starting extraction at {self.current_datetime}")
            print(f"User: {self.user_name}")
            print(f"Total kinases to process: {len(kinases)}")

            for idx, kinase in enumerate(kinases, 1):
                print(f"\nProcessing kinase {idx}/{len(kinases)}: {kinase}")
                
                try:
                    if self.search_pubmed(kinase):
                        papers = self.get_all_abstracts(kinase)
                        if papers:
                            result = {
                                'Kinase': kinase,
                                'Title': papers[0]['title'],
                                'Abstract': self.format_papers_for_excel(papers)
                            }
                            self.results.append(result)
                            print(f"Successfully extracted {len(papers)} abstracts for {kinase}")
                            
                            # Save partial results every 5 kinases
                            if idx % 5 == 0:
                                self.save_partial_results()
                        else:
                            print(f"No abstracts found for {kinase}")
                    else:
                        print(f"Failed to search for {kinase}")
                
                except Exception as e:
                    print(f"Error processing kinase {kinase}: {e}")
                    self.save_partial_results()
                    
                    # Try to recover the driver
                    try:
                        self.driver.quit()
                        time.sleep(5)
                        self.setup_driver()
                    except:
                        print("Failed to recover WebDriver. Saving progress and exiting...")
                        break

                time.sleep(Config.PAGE_LOAD_DELAY)

        except Exception as e:
            print(f"Critical error: {e}")
        finally:
            try:
                if self.results:
                    self.save_to_excel(self.results)
                    print(f"Final results saved to: {self.output_file_path}")
                if self.driver:
                    self.driver.quit()
            except Exception as e:
                print(f"Error during cleanup: {e}")

if __name__ == "__main__":
    try:
        print("Starting PubMed Kinase Scraper...")
        print(f"Current directory: {os.getcwd()}")
        print(f"Input file exists: {os.path.exists(Config.INPUT_FILE)}")
        
        scraper = PubMedKinaseScraper(
            input_file_path=Config.INPUT_FILE,
            output_file_path=Config.OUTPUT_FILE
        )
        scraper.process_kinases()
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        import traceback
        traceback.print_exc()