class Config:
    # File paths
    INPUT_FILE = "input.xlsx"
    OUTPUT_FILE = "output.xlsx"
    
    # User information
    USER_NAME = "IvinShajan"
    CURRENT_DATETIME = "2025-02-22 10:21:16"
    
    # Scraper settings
    MAX_PAPERS_PER_KINASE = 100
    WAIT_TIME = 20
    PAGE_LOAD_DELAY = 3
    MAX_RETRIES = 3
    CONNECTION_TIMEOUT = 30
    BATCH_SIZE = 10
    
    # Browser settings
    BROWSER_OPTIONS = [
        '--no-sandbox',
        '--disable-dev-shm-usage',
        '--disable-gpu',
        '--window-size=1920,1080',
        '--disable-features=TranslateUI',
        '--disable-extensions',
        '--disable-notifications',
        '--disable-infobars',
        '--disable-background-timer-throttling',
        '--disable-backgrounding-occluded-windows',
        '--disable-breakpad',
        '--disable-component-extensions-with-background-pages',
        '--disable-dev-shm-usage',
        '--disable-features=TranslateUI',
        '--disable-hang-monitor',
        '--disable-ipc-flooding-protection',
        '--disable-popup-blocking',
        '--disable-prompt-on-repost',
        '--disable-renderer-backgrounding',
        '--disable-sync',
        '--disable-translate',
        '--metrics-recording-only',
        '--no-first-run',
        '--safebrowsing-disable-auto-update',
        '--password-store=basic'
    ]

    # Human-related terms for filtering
    HUMAN_TERMS = [
        'human',
        'patient',
        'clinical',
        'person',
        'human cells',
        'human tissue',
        'human protein',
        'human study',
        'human sample',
        'human subject',
        'human experiment',
        'human trial',
        'human research',
        'human data',
        'human investigation',
        'human analysis',
        'human population',
        'human specimen',
        'human participant',
        'human volunteer'
    ]