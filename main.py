# TESTAR SERVIÇOS WEB


from pathlib import Path
from time import sleep

from myfunctions import load_excel_data, make_chrome_browser, open_webservices

# Path for project root
ROOT_FOLDER = Path(__file__).parent
# Path for dataset from excel
WORKBOOK_PATH = ROOT_FOLDER / 'servers.xlsx'


if __name__ == '__main__':
    TIME_TO_WAIT = 30
    options = ()
    # options = '--headless', '--disable-gpu'

    # Load dataset from Excel
    list_links = load_excel_data(WORKBOOK_PATH)

    # Create Chrome browser
    chrome_browser = make_chrome_browser(*options)

    # Open webservices in new tabs
    open_webservices(chrome_browser, *list_links)

    sleep(TIME_TO_WAIT)
    chrome_browser.quit()
