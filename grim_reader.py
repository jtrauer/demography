
from xlrd import open_workbook
import numpy
import matplotlib.pyplot as plt


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


def read_grim_sheet(workbook, sheet_name, years_to_keep=None, title_row_index=5, gender_row_index=3):
    """
    Function to read a single GRIM-formatted spreadsheet.

    Args:
        workbook: The entire spreadsheet
        sheet_name: String of the sheet of the book to be read
        years_to_keep: List of Booleans for the years of interest, only created when the all mortality sheet read
        title_row_index: Integer for the row with the titles in it
        gender_row_index: Integer for the row with the gender strings in it
    Returns:
        age_groups: List of the age groups strings
        years: List of the year integers
        genders: List of the strings for the genders
        final_array: The main array containing the data
        years_to_keep: Boolean structure corresponding to the years to be kept if it is all-cause mortality being read
    """

    # initialise
    sheet = workbook.sheet_by_name(sheet_name)
    data_type = []
    titles = {}
    working_array = {}
    gender = 'start'
    new_layer = False
    columns_to_ignore = ['', 'Total']

    for c in range(sheet.ncols):

        # update first level of dictionary indices if necessary, the "data type"
        if sheet.col_values(c)[gender_row_index] != u'':
            gender = str(sheet.col_values(c)[gender_row_index])
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
    genders = []
    final_array \
        = numpy.array(numpy.zeros(shape=(working_array['Persons'].shape[0], working_array['Persons'].shape[1], 0)))
    for gender in working_array:
        final_array = numpy.dstack((final_array, working_array[gender]))
        genders.append(gender)

    # if all data is zeros for that row across all layers, discard that row (year)
    if years_to_keep is None: years_to_keep = numpy.any(numpy.all(final_array, axis=0), axis=1)
    final_array = final_array[:, years_to_keep, :]
    years = list(numpy.array(years)[years_to_keep])

    return age_groups, years, genders, final_array, years_to_keep


def read_all_grim_sheets(sheet_names):
    """
    Master function loop over all sheets and read each one, then concatenate the sheets together along the fourth
    dimension.

    Args:
        sheet_names: The sheets that need to be read
    Returns:
         age_groups: Age group strings directly from the sheet reading function
         years: List of years as integers directly from teh sheet reading function
         final_array: The final data structure in four dimensions by age group, years, gender and sheet (cause of death)
    """

    # loop through sheet names
    for name in sheet_names:

        # ensure we keep the years from the all causes sheet, as there are several years with zero data in other sheets
        if name == sheet_names[0]:
            years_to_keep = None
        else:
            years_to_keep = years_to_keep

        # open excel workbook and find deaths sheet
        book = open_workbook('grim-' + name + '-2017.xlsx')

        # read with reading function above
        age_groups, years, genders, sheet_array, years_to_keep \
            = read_grim_sheet(book, 'Deaths', years_to_keep=years_to_keep)
        if name == sheet_names[0]: final_array = numpy.array(numpy.zeros(shape=list(sheet_array.shape) + [0L]))
        sheet_array = numpy.expand_dims(sheet_array, axis=3)

        # first dimension is age groups, second is years, third is gender, fourth is cause of death
        final_array = numpy.concatenate((final_array, sheet_array), axis=3)

    return age_groups, years, genders, final_array


if __name__ == '__main__':

    # specify spreadsheets to read and read them into single data structure - always put all-causes-combined first
    sheet_names = ['all-causes-combined']
    # 'all-certain-conditions-originating-in-the-perinatal-period',
    # 'all-certain-infectious-and-parasitic-diseases',
    # 'all-diseases-of-the-circulatory-system',
    # 'all-congenital-malformations-deformations-and-chromosomal-abnormalities',
    # 'all-diseases-of-the-blood-and-blood-forming-organs',
    # 'all-diseases-of-the-digestive-system',
    # 'all-diseases-of-the-ear-and-mastoid-process',
    # 'all-diseases-of-the-eye-and-adnexa',
    # 'all-diseases-of-the-genitourinary-system',
    # 'all-diseases-of-the-musculoskeletal-system-and-connective-tissue',
    # 'all-diseases-of-the-nervous-system',
    # 'all-diseases-of-the-respiratory-system',
    # 'all-diseases-of-the-skin-and-subcutaneous-tissue',
    # 'all-endocrine-nutritional-and-metabolic-diseases',
    # 'all-external-causes-of-morbidity-and-mortality',
    # 'all-mental-and-behavioural-disorders',
    # 'all-neoplasms',
    # 'all-pregnancy-childbirth-and-the-puerperium',
    # 'all-symptoms-signs-and-abnormal-clinical-and-laboratory-findings-not-elsewhere-classified',
    # 'asthma', 'breast-cancer', 'chronic-kidney-disease', 'colorectal-cancer',
    # 'chronic-obstructive-pulmonary-disease', 'coronary-heart-disease', 'diabetes', 'heart-failure',
    # 'lung-cancer', 'melanoma', 'osteoarthritis', 'osteoporosis', 'prostate-cancer', 'rheumatoid-arthritis',
    # 'stroke', 'cerebrovascular-disease', 'dementia-and-alzheimer-disease', 'hypertensive-disease',
    # 'kidney-failure', 'suicide', 'accidental-drowning', 'accidental-poisoning', 'assault',
    # 'land-transport-accidents', 'liver-disease']

    book = open_workbook('grim-' + sheet_names[0] + '-2017.xlsx')
    population_age_groups, population_years, genders, population_array, _ \
        = read_grim_sheet(book, 'Populations', title_row_index=14, gender_row_index=12)
    age_groups, years, genders, final_array = read_all_grim_sheets(sheet_names)

    # # quick example plot - absolute numbers
    # figure = plt.figure()
    # ax = figure.add_axes([0.1, 0.1, 0.6, 0.75])
    # for i in range(len(age_groups)):
    #     ax.plot(years, final_array[i, :, 0, :], label=age_groups[i])
    # handles, labels = ax.get_legend_handles_labels()
    # leg = ax.legend(handles, labels, bbox_to_anchor=(1.05, 1), loc=2, borderaxespad=0., frameon=False, prop={'size': 7})
    # figure.savefig('test_figure')

