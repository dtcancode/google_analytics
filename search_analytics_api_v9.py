import argparse
import sys
import time
import datetime
import xlsxwriter
from googleapiclient import sample_tools

# Declare arguments in command line
argparser = argparse.ArgumentParser(add_help=False)
argparser.add_argument('property_uri', type=str,
                       help=('Site or app URI to query data for (including '
                             'trailing slash).'))
argparser.add_argument('start_date', type=str,
                       help=('Start date of the requested date range in '
                             'YYYY-MM-DD format.'))
argparser.add_argument('end_date', type=str,
                       help=('End date of the requested date range in '
                             'YYYY-MM-DD format.'))

# Request API data
def execute_request(service, property_uri, request):
    print('                       ')
    return service.searchanalytics().query(
                siteUrl=property_uri, body=request).execute()

# Main function
page_query_list = []
row_ws1 = 2
def main(argv):
    service, flags = sample_tools.init(
      argv, 'webmasters', 'v3', __doc__, __file__, parents=[argparser],
      scope='https://www.googleapis.com/auth/webmasters.readonly')

    # prepare XLSX file
    title = 'pages_and_queries_'
    dates = str(flags.start_date) + '_' + str(flags.end_date)
    fileName = title + dates

    workbook = xlsxwriter.Workbook(fileName + '.xlsx', {'strings_to_urls': False})

    worksheet1 = workbook.add_worksheet('Pages and Queries')
    worksheet1.set_column('A:A', 35), worksheet1.set_column('B:B', 25), worksheet1.set_column('C:C', 8)
    worksheet1.set_column('D:D', 10), worksheet1.set_column('E:E', 8), worksheet1.set_column('F:F', 8)

    worksheet2 = workbook.add_worksheet('Totals, Date Range')
    worksheet2.set_column('A:A', 10), worksheet2.set_column('B:B', 10), worksheet2.set_column('C:C', 10)
    worksheet2.set_column('D:D', 10), worksheet2.set_column('E:E', 10), worksheet2.set_column('F:F', 10)

    worksheet3 = workbook.add_worksheet('Totals, Each Date')
    worksheet3.set_column('A:A', 10), worksheet3.set_column('B:B', 10), worksheet3.set_column('C:C', 10)
    worksheet3.set_column('D:D', 10), worksheet3.set_column('E:E', 10), worksheet3.set_column('F:F', 10)

    worksheet4 = workbook.add_worksheet('Totals, Each Page')
    worksheet4.set_column('A:A', 35), worksheet4.set_column('B:B', 10), worksheet4.set_column('C:C', 10)
    worksheet4.set_column('D:D', 10), worksheet4.set_column('E:E', 10)

    # Get pages and queries, worksheet1
    headings_ws1 = ['Page', 'Keywords', 'Clicks', 'Impressions', 'CTR', 'Position']
    worksheet1.write_row('A2', headings_ws1)
    worksheet1.autofilter(1, 0, 1000000, 5)
    def get_query_data(start_row):
        col_ws1 = 0
        global row_ws1
        global page_query_list

        request_queries = {
            'startDate': flags.start_date,
            'endDate': flags.end_date,
            'dimensions': ['page', 'query'],
            'rowLimit': 5000,
            'startRow': start_row
            }

        response_queries = execute_request(service, flags.property_uri, request_queries)
        if response_queries == {'responseAggregationType': 'byPage'}:
            for item in page_query_list:
                worksheet1.write_row(row_ws1, col_ws1, item)
                row_ws1 += 1
            workbook.close()
            sys.exit()

        request_query_data = response_queries['rows']
        print(str(len(request_query_data)) + ' Pages and Queries:')
        print('page', '|', 'keywords', '|', 'clicks', '|', 'impressions', '|', 'ctr', '|', 'position')

        for row in request_query_data:
            print(row['keys'][0], '|', row['keys'][1], '|', row['clicks'], '|', row['impressions'], '|',
                  round(row['ctr'], 3), '|', round(row['position'], 3))

            page_query_list.append([row['keys'][0], row['keys'][1], row['clicks'], row['impressions'],
                  round(row['ctr'], 3),round(row['position'], 3)])

        start_row += 5001
        time.sleep(1.5)
        get_query_data(start_row)

    # Get totals for the date range, worksheet2
    request_all_dates = {
        'startDate': flags.start_date,
        'endDate': flags.end_date}
    response_all_dates = execute_request(service, flags.property_uri, request_all_dates)
    query_data = response_all_dates['rows']
    title_all_dates = 'Totals for Date Range'
    print(title_all_dates)
    print('start date', '|', 'end date' , '|', 'clicks', '|', 'impressions', '|', 'ctr', '|', 'position')

    headings_ws2 = ['Start Date', 'End Date', 'Clicks', 'Impressions', 'CTR', 'Position']
    worksheet2.write_row('A2', headings_ws2)
    row_ws2 = 2
    col_ws2 = 0

    for row in query_data:
        print(flags.start_date, '|', flags.end_date, '|', row['clicks'], '|', row['impressions'], '|',
            round(row['ctr'],3), '|', round(row['position'],3))
        worksheet2.write_row(row_ws2, col_ws2, [flags.start_date, flags.end_date, row['clicks'], row['impressions'],
            round(row['ctr'], 3), round(row['position'], 3)])

    # Get totals for each date, worksheet3
    request_each_date = {
        'startDate': flags.start_date,
        'endDate': flags.end_date,
        'dimensions': ['date']
        }
    response_each_date = execute_request(service, flags.property_uri, request_each_date)
    query_data = response_each_date['rows']
    title_each_date = 'Totals for Each Date'
    print(title_each_date)
    print('date', '|', 'day', '|', 'clicks', '|', 'impressions', '|', 'ctr', '|', 'position')

    headings_ws3 = ['Date', 'Day', 'Clicks', 'Impressions', 'CTR', 'Position']
    worksheet3.write_row('A2', headings_ws3)
    worksheet3.autofilter(1, 0, 1000000, 5)
    row_ws3 = 2
    col_ws3 = 0

    for row in query_data:
        day_of_week = datetime.datetime.strptime(str(row['keys'][0]), "%Y-%m-%d").strftime('%A')
        print(row['keys'][0], '|', day_of_week, '|', row['clicks'], '|', row['impressions'], '|',
            round(row['ctr'],3), '|', round(row['position'],3))
        worksheet3.write_row(row_ws3, col_ws3, [row['keys'][0], day_of_week, row['clicks'], row['impressions'],
            round(row['ctr'], 3), round(row['position'], 3)])
        row_ws3 += 1

    # Get totals for each page, worksheet4
    request_each_page = {
        'startDate': flags.start_date,
        'endDate': flags.end_date,
        'dimensions': ['page']
        }
    response_each_page = execute_request(service, flags.property_uri, request_each_page)
    query_data = response_each_page['rows']
    title_each_page = 'Totals for Each Page'
    print(title_each_page)
    print('page', '|', 'clicks', '|', 'impressions', '|', 'ctr', '|', 'position')

    headings_ws4 = ['Page', 'Clicks', 'Impressions', 'CTR', 'Position']
    worksheet4.write_row('A2', headings_ws4)
    worksheet4.autofilter(1, 0, 1000000, 4)
    row_ws4 = 2
    col_ws4 = 0

    for row in query_data:
        print(row['keys'][0], '|', row['clicks'], '|', row['impressions'], '|',
            round(row['ctr'],3), '|', round(row['position'],3))
        worksheet4.write_row(row_ws4, col_ws4, [row['keys'][0], row['clicks'], row['impressions'],
            round(row['ctr'], 3), round(row['position'], 3)])
        row_ws4 += 1

    get_query_data(0)

# Run program via command line
if __name__ == '__main__':
  main(sys.argv)
