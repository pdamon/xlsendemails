import xlwings as xw
import pandas as pd
from re import match

import dominate
from dominate.tags import *
from dominate.util import raw

import win32com as win32


def process_news():
    wb = xw.Book.caller()
    # wb = xw.Book('../researchpyxll3.xlsm')

    # get data
    sht = wb.sheets['RAW DATA']
    data = sht.range('A1:J1000').options(pd.DataFrame, index=False).value

    # get interest lists
    sht = wb.sheets['INTEREST LIST']
    il = sht.range('A1:ZZ1000').options(pd.DataFrame, index=False).value.drop(labels=[None], axis=1)

    prev = 'INTEREST LIST'
    for client in il.columns:
        try:
            wb.sheets.add(name=client, after=prev)

        except ValueError:
            # this occurs when the sheet already exists
            pass

        df = filter_dataframe(il[client].dropna(), data)

        if len(df) > 0:
            df = df.sort_values(['Date'])
            df['Link'] = df['Bloomberg Link'].apply(lambda x: 'bbg://news/stories/{}'.format(x.split(" ")[1]))
            # df['hl_int'] = df['Headline']
            # df['Headline'] = df.apply(lambda x: '=HYPERLINK("{}", "{}")'.format('bbg://news/stories/{}'.format(x['Bloomberg Link'].split(" ")[1]), x['Headline']), axis=1)

            # df = df[['Primary Tickers', 'Headline', 'Broker', 'Bloomberg Link', 'Link', 'hl_int']]
            df = df[['Primary Tickers', 'Headline', 'Broker', 'Bloomberg Link', 'Link']]

            df = df.sort_values(["Primary Tickers", "Broker"])

            sht = wb.sheets[client]

            sht.range('A1').options(index=False).value = df
            sht.autofit()

        prev = client


def str_to_series(x):
    try:
        spl = pd.Series(x.split(","))

        ss = spl.apply(lambda x: x.rstrip().lstrip())

        return ss
    except AttributeError:
        return pd.Series(dtype=object)


def filter_dataframe(c, d):
    # d = d.drop_duplicates(['Primary Tickers', 'Secondary Tickers', 'Headline', 'Broker', 'Action', 'Rating', 'Pg', 'Content Type'])
    d = d.drop_duplicates(['Primary Tickers', 'Headline'])

    df = pd.DataFrame()
    for _, r in d.iterrows():
        pt = str_to_series(r['Primary Tickers'])

        if pt.isin(c).any() and not match(r"\[Delayed\].*", r['Headline']):
            df = df.append(r)

    return df


def make_subject_line(news):
    sl = "Interest List Broker Notes:"

    s = pd.Series()
    for i in news['Primary Tickers']:
        s = s.append(str_to_series(i))

    s = s.drop_duplicates()

    for i in s:
        sl += " " + i + ","

    return sl[:-1]


def send_emails():
    wb = xw.Book.caller()
    # wb = xw.Book("../researchpyxll3.xlsm")
    sht = wb.sheets['Email List']

    clients = sht.range('A1:C1000').options(pd.DataFrame, index=False).value.dropna()

    for _, c in clients.iterrows():
        create_email(wb, c['Client'], c['Address'], c['Send'])


def create_email(wb, client, addr, send):
    # get the news for the client
    sht = wb.sheets[client]
    data = sht.range('A1:E1000').options(pd.DataFrame, index=False).value.dropna()

    # get the client's watchlist
    sht = wb.sheets['INTEREST LIST']
    il = sht.range('A1:ZZ1000').options(pd.DataFrame, index=False).value.drop(labels=[None], axis=1)
    wl = il[client].dropna()

    # create the subject line
    subject = make_subject_line(data)

    # create email HTML
    html = make_email(data, wl, "", "", pd.Series(), pd.Series())

    # open email in outlook
    outlook = win32.client.gencache.EnsureDispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.To = addr
    mail.Subject = subject
    mail.HTMLBody = html

    if send == 'YES DO SEND THIS':
        mail.Send()
    else:
        mail.Display()


def make_email(news, watchlist, name, comments, pictures, pic_comments):
    """
    Args:
        news: dataframe with relevant broker news
        watchlist: series of watchlist
        name: the name of the person the email is sent to
        comments: str with general
        pictures: the path of pictures to be included
        pic_comments: the path of the picture comments to be included

    Returns:
        A string with the html to embed in the email
    """
    assert(len(pictures) == len(pic_comments))
    # make the document
    doc = dominate.document(title='Email')

    # *{
    #     font - family: Arial, Helvetica, sans - serif !important;
    # color: black;
    # }

    with doc.head:
        raw("""
        <style>
        * {
            font-family: Arial, Helvetica, sans-serif !important;
            color: black;
        }
        table,
        th,
        td {
          padding: 2px;
          border: 1px solid black;
          border-collapse: collapse;
        }
      </style>
        """)

    with doc:
        # make the title area
        with div():
            h1(name)
            p(comments)

        # make the market news
        with div():
            h4("Interest List Broker Notes")
            with table() as t:
                my_tr = tr()
                # make the header row
                for hl in ['Primary Tickers', 'Headline', 'Broker', 'Bloomberg Link']:
                    my_tr.add(th(hl))

                t.add(my_tr)

                # insert the news data
                for _, row in news.iterrows():
                    my_tr = tr()
                    my_tr.add(td(row[0]))
                    my_tr.add(td(a(row[1], href=row[4])))
                    my_tr.add(td(row[2]))
                    my_tr.add(td(row[3]))

                    t.add(my_tr)

        # put some cool graphs or something here
        if len(pictures) > 0:
            with div():
                h4("Some other title")
                with table() as t:
                    for j in range(len(pictures)):
                        my_tr = tr()
                        my_tr.add(td(pictures[j]))
                        my_tr.add(td(pic_comments[j]))

                        t.add(my_tr)

        # make the interest list
        with div():
            h4("Your Interest List")

            with table():
                for i in watchlist:
                    tr().add(td(i))

    return str(doc)[15:]


if __name__ == "__main__":
    # process_news()
    # create_email("APD", "parker.damon32@gmail.com")
    send_emails()
