
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
        values_found: The list after processing as a list of integers
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
        title = remove_element_from_unicode(sheet.col_values(c)[title_row_index], 8211, u' to ')

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


def convert_grim_string(string_to_convert):
    """
    Just a function to access a dictionary of string conversions. Will gradually build out as we need to for outputting.

    Args:
        string_to_convert: The raw input string that isn't nicely formatted
    Returns:
        The converted string
    """

    conversion_dictionary \
        = {'all-external-causes-of-morbidity-and-mortality': 'External causes',
           'all-diseases-of-the-circulatory-system': 'Circulatory diseases',
           'all-neoplasms': 'Neoplasms'}

    if string_to_convert in conversion_dictionary:
        return conversion_dictionary[string_to_convert]
    else:
        return string_to_convert


def distribute_missing_across_agegroups(final_array, age_groups):
    """
    Distribute the missing data proportionately across age groups. Note that is typically less than 0.1% of all data,
    but probably still better to do to improve the sense of the absolute rates of death.

    Args:
        final_array: The final data array
        age_groups: List of age groups, so that the Missing one can be indexed (although it's always the last one)
    Returns:
        adjusted_for_missing_array: Array structured as final_array was, but with no missing column and adjusted age
            group values
    """

    adjusted_for_missing_array \
        = numpy.zeros([final_array.shape[0] - 1, final_array.shape[1], final_array.shape[2], final_array.shape[3]])
    for y in range(final_array.shape[1]):
        for g in range(final_array.shape[2]):
            for c in range(final_array.shape[3]):
                prop_missing = final_array[age_groups.index('Missing'), y, g, c] \
                               / sum(final_array[:age_groups.index('Missing'), y, g, c])
                adjusted_for_missing_array[:, y, g, c] \
                    = final_array[:age_groups.index('Missing'), y, g, c] * (1. + prop_missing)
    return adjusted_for_missing_array


def find_rates_from_deaths_and_populations(death_array, pop_array, n_sheets):
    """
    Divides the matrix of numbers of deaths by the population matrix.

    Args:
        death_array: Array of deaths, which should be adjusted such that "Missing" age category isn't present
        pop_array: Array of total population numbers to be used as denominator
        n_sheets: The number of spreadsheets read in to apply this function to
    Returns:
        rates_array: The array of death rates per year
    """

    rates_array = numpy.zeros_like(death_array)
    for s in range(n_sheets): rates_array[:, :, :, s] = numpy.divide(death_array[:, :, :, s], pop_array)
    return rates_array


def restrict_population_to_relevant_years(pop_array, data_years):
    """
    Restrict the population array (which comes from the GRIM data with more years than the death data come with) to the
    years that are relevant to the death data being read in.

    Args:
        pop_array: The full, unrestricted population array
        data_years: The years that are applicable from the death data array
    """

    return pop_array[:, population_years.index(data_years[0]):population_years.index(data_years[-1]) + 1, :]


def find_string_from_dict(string, capitalise=True):

    string_dictionary = {'all-diseases-of-the-circulatory-system': 'cardiovascular disease',
                         'all-neoplasms': 'cancer'}
    string_to_return = string_dictionary[string] if string in string_dictionary else string
    if capitalise:
        return string_to_return[0].upper() + string_to_return[1:]
    else:
        return string_to_return


if __name__ == '__main__':

    # first dimension is age groups, second is years, third is gender, fourth is cause of death

    # specify spreadsheets to read and read them into single data structure - always put all-causes-combined first
    sheet_names = ['all-causes-combined',
                   'all-diseases-of-the-circulatory-system',
                   'all-neoplasms']

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

    # read population totals
    book = open_workbook('grim-' + sheet_names[0] + '-2017.xlsx')
    population_age_groups, population_years, genders, population_array, _ \
        = read_grim_sheet(book, 'Populations', title_row_index=14, gender_row_index=12)

    # read death data spreadsheets
    age_groups, years, genders, final_array = read_all_grim_sheets(sheet_names)

    # specify time range
    start_year = 1907
    finish_year = 2014

    # adjust for missing data, restrict population array to relevant years and calculate rates
    adjusted_array = distribute_missing_across_agegroups(final_array, age_groups)
    population_array_relevant_years = restrict_population_to_relevant_years(population_array, years)
    rates = find_rates_from_deaths_and_populations(adjusted_array, population_array_relevant_years, len(sheet_names))

    # # create graph of total death rates by age groups over time
    # for gender in genders:
    #
    #     # quick example plot - mortality rates by age group
    #     figure = plt.figure()
    #     ax = figure.add_axes([0.1, 0.1, 0.6, 0.75])
    #     for i in range(len(age_groups) - 1):
    #         ax.plot(range(years.index(start_year), years.index(finish_year) + 1),
    #                 rates[i, years.index(start_year):years.index(finish_year) + 1,
    #                 genders.index(gender), sheet_names.index('all-causes-combined')], label=age_groups[i])
    #     handles, labels = ax.get_legend_handles_labels()
    #     leg = ax.legend(handles, labels, bbox_to_anchor=(1.05, 1), loc=2, borderaxespad=0., frameon=False,
    #                     prop={'size': 7})
    #     ax.set_title(gender)
    #     ax.set_ylim((0., 0.35))
    #     plt.setp(ax.get_xticklabels(), fontsize=10)
    #     plt.setp(ax.get_yticklabels(), fontsize=10)
    #     figure.savefig('mortality_figure_' + gender)
    #
    # # deaths by cause with limitation by age group
    # for upper_age_limit in ['70 to 74', '75 to 79']:
    #     denominators \
    #         = numpy.sum(population_array[:, population_years.index(start_year):population_years.index(finish_year) + 1,
    #                     genders.index('Persons')], axis=0)
    #     numerators = {}
    #     rates = {}
    #     causes = ['all-causes-combined', 'all-diseases-of-the-circulatory-system', 'all-neoplasms']
    #     for cause in causes:
    #         numerators[cause] = numpy.sum(adjusted_array[:age_groups.index(upper_age_limit),
    #                                       years.index(start_year):years.index(finish_year) + 1,
    #                                       genders.index('Persons'), sheet_names.index(cause)], axis=0)
    #         rates[cause] = [i / j for i, j in zip(numerators[cause], denominators)]
    #
    #     figure = plt.figure()
    #     ax = figure.add_axes([0.1, 0.1, 0.6, 0.75])
    #     for cause in causes:
    #         ax.plot(years[years.index(start_year):years.index(finish_year) + 1],
    #                 rates[cause], label=convert_grim_string(cause))
    #     handles, labels = ax.get_legend_handles_labels()
    #     leg = ax.legend(handles, labels, bbox_to_anchor=(1.05, 1), loc=2, borderaxespad=0., frameon=False,
    #                     prop={'size': 7})
    #     ax.set_title('Death rates by cause for under ' + upper_age_limit[:2] + 's')
    #     ax.set_ylim((0., 3e-3))
    #     ax.set_xlabel('Year', fontsize=10)
    #     ax.set_ylabel('Rate per capita per year', fontsize=10)
    #     plt.setp(ax.get_xticklabels(), fontsize=10)
    #     plt.setp(ax.get_yticklabels(), fontsize=10)
    #     figure.savefig('mortality_figure_cause_under ' + upper_age_limit[:2] + 's')

    figure = plt.figure()
    life_tables = {}
    cumulative_deaths_by_cause = {}
    integer_ages = []

    # construct life tables and cumulative death structures for each calendar year
    for year in range(start_year, finish_year):

        # the life table list and the running value to populate it
        survival_total = 1.
        life_tables[year] = [1.]

        # the cumulative death structures and the value to populate it, by cause of death
        cumulative_deaths_by_cause[year] = {}
        for cause in sheet_names:
            cumulative_deaths_by_cause[year][cause] = [0.]
            cumulative_deaths = 0.

            # looping over each age group
            for age_group in range(rates.shape[0]):

                # find the applicable rate
                rate = rates[age_group, years.index(year), genders.index('Persons'), sheet_names.index(cause)]

                # applying it for each individual age in years
                for i in range(5):
                    if cause == 'all-causes-combined':
                        survival_total *= 1. - rate
                        life_tables[year].append(survival_total)
                    integer_age = age_group * 5 + i
                    if year == start_year and cause == sheet_names[0]:
                        integer_ages.append(integer_age)
                    cumulative_deaths += life_tables[year][integer_age] * rate
                    cumulative_deaths_by_cause[year][cause].append(cumulative_deaths)

    # plot cumulative survival graphs by year and age
    n_plots, rows, columns, base_font_size = 5, 2, 3, 8
    plt.style.use('ggplot')
    for n_plot in range(n_plots):
        year = 2025 + n_plot * 15 - n_plots * 15
        ax = figure.add_subplot(rows, columns, n_plot + 1)
        stacked_data = {'base': numpy.zeros(len(life_tables[year])),
                        'survival': life_tables[year],
                        'other causes': numpy.ones(len(life_tables[year]))}
        ordered_list_of_stacks = ['base', 'survival']
        new_data = life_tables[year]
        for cause in cumulative_deaths_by_cause[year]:
            if cause != 'all-causes-combined':
                new_data = [i + j for i, j in zip(new_data, cumulative_deaths_by_cause[year][cause])]
                stacked_data[cause] = new_data
                ordered_list_of_stacks.append(cause)
        ordered_list_of_stacks.append('other causes')
        for i in range(1, len(ordered_list_of_stacks)):
            ax.fill_between(integer_ages,
                            stacked_data[ordered_list_of_stacks[i - 1]][:-1],
                            stacked_data[ordered_list_of_stacks[i]][:-1],
                            color=list(plt.rcParams['axes.prop_cycle'])[i - 1]['color'],
                            label=find_string_from_dict(ordered_list_of_stacks[i]))
        handles, labels = ax.get_legend_handles_labels()
        if n_plot >= columns: ax.set_xlabel('Age', fontsize=base_font_size)
        if n_plot % columns == 0: ax.set_ylabel('Proportion', fontsize=base_font_size)
        if n_plot == n_plots - 1:
            ax.legend(handles, labels, bbox_to_anchor=(1.2, 1), loc=2, frameon=False, prop={'size': 9})
        plt.setp(ax.get_xticklabels(), fontsize=base_font_size - 2)
        plt.setp(ax.get_yticklabels(), fontsize=base_font_size - 2)
        ax.set_title(year, fontsize=base_font_size + 2)
        ax.set_xlim((50., 89.))
    figure.savefig('lifetable')




