
from xlrd import open_workbook


def remove_element_from_unicode(unicode_string, element_number, insertion_element):
    """
    Removes a specific unicode character from a unicode string and replaces it with a unicode string.
    Then returns the string with the replaced value as plain string.
    Unbelievably, I couldn't find a simpler way to do this that only took one line.
    """

    assert type(insertion_element) == unicode, 'Insertion element not unicode'
    assert type(element_number) == int, 'Unicode code number not integer'
    revised_string = []
    for i in range(len(unicode_string)):
        if ord(unicode_string[i]) == element_number:
            revised_string.append(insertion_element)
        else:
            revised_string.append(unicode_string[i])
    return str(''.join(revised_string))


def convert_to_integer_if_possible(values_list):
    """
    Another function that I can't believe I need. Takes a list that includes mostly integers values in a different
    format (e.g. float) that can be converted easily into integers and other rubbish and returns all as integers.

    Args:
        values_list: The list to be interrogated
    Returns:
        valuse_found: The list after processing as a list of integers
    """

    values_found = []
    for i in range(len(values_list)):
        try:
            values_found.append(int(values_list[i]))
        except:
            values_found.append(0)
    return values_found


# open excel workbook and find deaths sheet
book = open_workbook('grim-all-causes-combined-2017.xlsx')
sheet = book.sheet_by_name('Deaths')

# initialise
data_type = ''
data = {data_type: {}}
for c in range(sheet.ncols):

    # update first level of dictionary indices if necessary, the "data type"
    if sheet.col_values(c)[3] != u'':
        data_type = str(sheet.col_values(c)[3])
        data[data_type] = {}

    # find second level of dictionary indices, the "title"
    title_row_index = 5
    title = remove_element_from_unicode(sheet.col_values(c)[title_row_index], 8211, u'to')

    # process column values
    column_values = convert_to_integer_if_possible(sheet.col_values(c)[title_row_index:])
    data[data_type][title] = column_values



# code for pandas (incomplete and not working as it should - have abandoned for now)
# import pandas
# grim_all_causes_file = pandas.ExcelFile('grim-all-causes-combined-2017.xlsx')
# grim_all_causes = {'male': grim_all_causes_file.parse('Deaths', header=5, usecols=range(0, 21))}
# grim_all_causes['male'].fillna(0, inplace=True)


