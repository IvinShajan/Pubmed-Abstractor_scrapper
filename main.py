from scraper import PubMedKinaseScraper
from config import Config

def main():
    scraper = PubMedKinaseScraper(
        input_file_path=Config.INPUT_FILE,
        output_file_path=Config.OUTPUT_FILE
    )
    scraper.process_kinases()

if __name__ == "__main__":
    main()