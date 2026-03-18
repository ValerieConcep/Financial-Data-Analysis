import datetime
import pprint
import statistics
import urllib
import openpyxl
import requests

# copy in getEndpoint function for FMP
def getEndpoint(endpoint, version, parameters):
    baseUrl = f"https://financialmodelingprep.com/{version}"

    endpointUrl = f"{baseUrl}/{endpoint}"

    parameters['apikey'] = "6nbL5vjOQHAxBFF293yE8OPyZsB9CAIV" # Copy your API key here
    headers = {}
    payload = {}

    print("Getting Endpoint: " + endpointUrl + "?" + urllib.parse.urlencode(parameters))
    response = requests.request("GET", endpointUrl, headers=headers, data=payload, params=parameters)
    response_data = response.json()
    return response_data
# create workbook for companies
stock_wb= openpyxl.Workbook()
ws1= stock_wb.active



# add worksheets for company data and stock price data
company_ws= stock_wb.create_sheet("Company Data")
stock_ws= stock_wb.create_sheet("Stock Data")

# add headers to company/stock price worksheets
company_ws.append([
    "Symbol", "Company Name", "Sector", "Exchange",
    "Volatility", "Trend Slope"
])

stock_ws.append([
    "Symbol", "Date", "Open", "Close"
])

# read in stocktwits file
with open ("stocklist.txt") as file_pointer:
    file_contents= file_pointer.read()

#print(len(file_contents.splitlines()))
# split file into list of symbols
ticker_list= []
for line_value in  (file_contents.splitlines()):
    user_lastname= "CD"
    if user_lastname in line_value:
        ticker_list= line_value[3:].split(",")

        break


print("Tickers:",ticker_list)


"stocklist.txt"


# loop through collection of symbols; for each symbol
#Loop through tickers
for ticker_index,ticker in enumerate(ticker_list):
    print("Processing:",ticker)

    # download profile data
    profile_request = getEndpoint(
        endpoint="profile",
        version="stable",
        parameters={"symbol": ticker}
    )


    company_profile= profile_request[0]

    company_name = company_profile.get("companyName")
    sector = company_profile.get("sector")
    exchange = company_profile.get("exchange")
    pprint.pprint(profile_request)

    #Stock Prices
    # download stock price data
    price_request = getEndpoint(
            endpoint="historical-price-full",
            version="stable",
            parameters={"symbol": ticker}
    )

    historical_price= price_request.get('historical-price-full')
    # create price trend dictionary of x and y values
    close_price= []
    x_values=[]
    y_values=[]
    # loop through collection of days of stock prices; for each day

    for index, day in enumerate(historical_price[:30]):
        # extract date, open, close, etc.
        date= day.get("date")
        open_price = day.get("open")
        close_price= day.get("close")


        # create 'record' array for stock price data points and append to worksheet
        stock_ws.append([ticker, date, open_price, close_price])
        # append price to y values and day number to x values
        close_price.append(close_price)
        x_values.append(index)
        y_values.append(close_price)

        # calculate volatility

    votality= statistics.stdev(close_price) if len(close_price)> 1  else 0

    # calculate price trend slope
    if len(x_values) > 1:
        slope = (y_values[-1] - y_values[0]) / (x_values[-1] - x_values[0])
    else:
        slope = 0

    # create a 'record' array to hold company data points
    company_ws.append([
        ticker,
        company_name,
        sector,
        exchange,
        round(votality, 2),
        round(slope, 2)
    ])
    # append record to appropriate worksheet

# save workbook
stock_wb.save("project3.xlsx")
