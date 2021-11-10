import pandas

if __name__ == '__main__':

    # Parse Excel File to extract the column "email", drop duplicates and drop empty
    parse_excel = pandas.read_excel(io="Lista.xls", usecols=["email"]).drop_duplicates(keep='last').dropna()

    # Print each element
    for emails in parse_excel['email']:
        print(emails)
