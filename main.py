'''
    This program asks the user to enter a stock ticker and then displays the latests news for that ticker.
'''

import requests
import pandas as pd

from bs4 import BeautifulSoup
from datetime import datetime

loading = False;

def main():
    # show welcome message
    print("Welcome to the Stock Scraper!");
    print("Please input tickers seperated by a space (e.g. AAPL MSFT GOOG) \n");

    # ask for the ticker
    ticker = input("Enter stock tickers: ")

    # add loading message
    print("Loading ...")

    all_tickers = ticker.split(" ");
    all_urls = getURLs(all_tickers);

    all_requests = getAllRequests(all_urls);

    export_data = pd.DataFrame(columns = ["Ticker", "Title", "Link"])

    counter = 0;

    # parse the requests
    for r in all_requests:
        soup = BeautifulSoup(r.text, "html.parser")

        content = soup.find_all("div", class_="article__content")
        all_news = []

        for n in content:
            header = n.find("h3", class_="article__headline")
            detail = n.find("div", class_="article__details")

            try:
                title = header.find("a", class_="link")
                link = header.find("a", class_="link").get("href")

                # check if link is valid
                if "http" in link:
                    title = title.get_text().strip()
                    link = link.strip()

                    # create dataframe and add to bottom of dataframe
                    export_data.loc[export_data.shape[0]] = [all_tickers[counter].upper(), title, link]
            except:
                pass

        counter += 1;

    exportToExcel(export_data)
    print("\nDone!")


# get all urls for the tickers
def getURLs(tickers):
    all_urls = []

    for ticker in tickers:
        url = "https://www.marketwatch.com/investing/stock/" + ticker
        all_urls.append(url)

    return all_urls

# get all requests for the urls and return a list of requests, if there is an error, exit the program and print the error
def getAllRequests(urls):
    all_requests = []

    for url in urls:
        r = requests.get(url)

        # check if url has substring "https://www.marketwatch.com/search?"
        if "https://www.marketwatch.com/search?" in r.url:
            print("Error - invalid ticker")
            exit()
        else:
            all_requests.append(r)

    return all_requests

# exports the dataframe to an excel file with the current date as the sheet name
def exportToExcel(dataframe):
    today = datetime.today().strftime('%Y-%m-%d')

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('reading_list.xlsx', engine = 'xlsxwriter')

    # Write the dataframe data to XlsxWriter. Turn off the default header and
    # index and skip one row to allow us to insert a user defined header.
    dataframe.to_excel(writer, sheet_name=today, startrow = 1, header = False, index = False)

    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets[today]

    # Get the dimensions of the dataframe.
    (max_row, max_col) = dataframe.shape

    # Create a list of column headers, to use in add_table().
    column_settings = [{'header': column} for column in dataframe.columns]

    # Add the Excel table structure. Pandas will add the data.
    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

    # Make the columns wider for clarity.
    worksheet.set_column(0, max_col - 1, 12)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()


if __name__ == '__main__':
    main()