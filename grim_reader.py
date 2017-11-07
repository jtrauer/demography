
from xlrd import open_workbook
import numpy

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
data_type = []
title_row_index = 5
titles = {}
working_array = {}
gender = 'start'
new_layer = False
columns_to_ignore = ['', 'Total']

for c in range(sheet.ncols):

    # update first level of dictionary indices if necessary, the "data type"
    if sheet.col_values(c)[3] != u'':
        gender = str(sheet.col_values(c)[3])
        new_layer = True
    data_type.append(gender)

    # find second level of dictionary indices, i.e. the "title"
    title = remove_element_from_unicode(sheet.col_values(c)[title_row_index], 8211, u'to')

    # process column values
    column_values = convert_to_integer_if_possible(sheet.col_values(c)[title_row_index:])

    # record year values
    if 'Year' in title:
        years = column_values

    # ignore columns with no data
    elif title in columns_to_ignore:
        pass

    # create main data structure
    elif new_layer:
        working_array[gender] = numpy.array(column_values)
        new_layer = False
        titles[gender] = [title]
    elif gender != 'start':
        working_array[gender] = numpy.vstack((working_array[gender], numpy.array(column_values)))
        titles[gender].append(title)

# depth stack the arrays created above
age_groups = titles['Persons']
layer_titles = []
final_array = numpy.array(numpy.zeros(shape=(working_array['Persons'].shape[0], working_array['Persons'].shape[1], 0)))
for gender in working_array:
    final_array = numpy.dstack((final_array, working_array[gender]))
    layer_titles.append(gender)

# if all data is zeros for that row across all layers, discard that row (year)
years_to_keep = numpy.any(numpy.all(final_array, axis=0), axis=1)
final_array = final_array[:, years_to_keep, :]
years = list(numpy.array(years)[years_to_keep])

# first dimension is age groups, second is years, third is gender for the final_array





# code for pandas (incomplete and not working as it should - have abandoned for now)
# import pandas
# grim_all_causes_file = pandas.ExcelFile('grim-all-causes-combined-2017.xlsx')
# grim_all_causes = {'male': grim_all_causes_file.parse('Deaths', header=5, usecols=range(0, 21))}
# grim_all_causes['male'].fillna(0, inplace=True)


