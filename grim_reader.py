
from xlrd import open_workbook
import numpy
import matplotlib.pyplot as plt


''' static methods'''


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


def read_standard_population():
    """
    Read in Australian Standard Population 2001 from ABS.
    """

    book = open_workbook('australian_standard_population_2001.xls')
    sheet = book.sheet_by_name('Table_1')
    data, titles = {}, sheet.row_values(5)
    for i in range(sheet.nrows):
        if sheet.col_values(titles.index(u'Persons'))[i] not in [u'', u'Persons']:
            if sheet.col_values(titles.index(u'Age (years)'))[i] == u'100 and over':
                age_key = 100
            elif sheet.col_values(titles.index(u'Age (years)'))[i] == u'Total':
                age_key = 'Total'
            else:
                age_key = int(sheet.col_values(titles.index(u'Age (years)'))[i])
            data[age_key] = int(sheet.col_values(titles.index(u'Persons'))[i])
    return data


def exclude_non_integer_elements_from_dict(dictionary):
    """
    Exclude any elements from a dictionary that do not have an integer as their key.

    Args:
        dictionary: The original dictionary that may have a mix of integer and non-integer keys
    Returns:
        Dictionary with only integer keys
    """

    return {i: dictionary[i] for i in dictionary if type(i) == int}


def sum_dict_over_brackets(dictionary, bracket_size=5):
    """
    Sums all the values within a regular range of integer values referring to the keys of the dictionary being analysed.

    Args:
        dictionary: The dictionary to be summed
        bracket_size: Integer for the width of the brackets to sum the keys over
    Returns:
        summed_dictionary: A new dictionary with keys the lower limits of the summations
    """

    revised_dict = exclude_non_integer_elements_from_dict(dictionary)

    # initialise
    summed_dictionary, within_bracket_n = {0: 0}, 0

    # cycle through all integer values that could be present in dictionary
    for i in range(max(revised_dict.keys()) + 1):

        # skip on and create new key to summed dictionary once finished the previous key
        if i % bracket_size == 0 and i > 0:
            within_bracket_n += bracket_size
            summed_dictionary[within_bracket_n] = 0

        # increment dictionary
        summed_dictionary[within_bracket_n] += revised_dict[i]

    # return summed dictionary
    return summed_dictionary


def sum_last_elements_of_dict(dictionary, value=85):
    """
    Aggregate the upper values of an integer-keyed dictionary into one.

    Args:
        dictionary: The dictionary that needs to be summed
        value: The highest value of the revised dictionary, to put all of the upper groups into
    Returns:
        dictionary: with upper values summed into the top key
    """

    revised_dict = exclude_non_integer_elements_from_dict(dictionary)
    assert value in revised_dict, 'Requested value for summing not found in dictionary'
    for k in revised_dict.keys():
        if k > value:
            revised_dict[value] += revised_dict[k]
            del revised_dict[k]
    return revised_dict


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
    if years_to_keep is None:
        years_to_keep = numpy.any(numpy.all(final_array, axis=0), axis=1)
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
    for n, name in enumerate(sheet_names):

        # ensure we keep the years from the all causes sheet, as there are several years with zero data in other sheets
        if n == 0:
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
           'all-diseases-of-the-circulatory-system': 'Cardiovascular disease',
           'all-neoplasms': 'Neoplasms',
           'all-causes-combined': 'All causes',
           'Persons': 'Both genders'}

    if string_to_convert in conversion_dictionary:
        return conversion_dictionary[string_to_convert]
    else:
        return string_to_convert[0].upper() + string_to_convert[1:].replace('-', ' ')


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
                prop_missing = 0. if numpy.isnan(prop_missing) else prop_missing
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
    for s in range(n_sheets):
        rates_array[:, :, :, s] = numpy.divide(death_array[:, :, :, s], pop_array)
    return rates_array


def restrict_population_to_relevant_years(pop_array, data_years, population_years):
    """
    Restrict the population array (which comes from the GRIM data with more years than the death data come with) to the
    years that are relevant to the death data being read in.

    Args:
        pop_array: The full, unrestricted population array
        data_years: The years that are applicable from the death data array
        population_years: The years available in the population matrix
    """

    return pop_array[:, population_years.index(data_years[0]):population_years.index(data_years[-1]) + 1, :]


def find_string_from_dict(string, capitalise=True):

    string_dictionary = {'all-diseases-of-the-circulatory-system': 'cumulative cardiovascular deaths',
                         'all-neoplasms': 'cumulative cancer deaths',
                         'all-causes-combined': 'all causes'}
    string_to_return = string_dictionary[string] if string in string_dictionary else string
    if capitalise:
        return string_to_return[0].upper() + string_to_return[1:]
    else:
        return string_to_return


def find_agegroup_values_from_strings(age_group_strings):
    """
    Function to extract the integer values of the age groups from their strings.

    Args:
        age_group_strings: The list containing the string descriptions of the age groups
    Returns:
        start_ages: The starting age for each age group
        end_ages: The ending age for each age group
    """

    start_ages = []
    end_ages = []
    for age_group in age_group_strings:
        age_group_splits = age_group.split(' ')
        if age_group_splits[0] == '85+':
            start_ages.append(85)
            end_ages.append(float('inf'))
        elif age_group_splits[0] == 'Missing':
            start_ages.append(0)
            end_ages.append(float('inf'))
        else:
            start_ages.append(int(age_group_splits[0]))
            end_ages.append(int(age_group_splits[-1]))
    return start_ages, end_ages


def karup_king_interpolation(group_index, within_group_index, last_age_group_index, data, age_group_width=5.):
    """
    Method to interpolate data to yearly intervals using the relatively simple Karup-King approach, which multiplies
    pre-defined coefficients by the rates in the age groups of interest and those on either side - except in the case
    where these are unavailable (i.e. the first or last) for which the closest three age groups are taken.

    Args:
        group_index: The index for the age group of interest
        within_group_index: Distance through the subgroup being analysed (i.e the five years)
        last_age_group_index: The index for the highest age group to be analysed
        data: List (or one-dimensional array) for the quantities being smoothed
        age_group_width: The number of single ages in the age group (currently has to be five)
    Returns:
        interpolated_estimate: The interpolated rate for the single year of interest
    """

    coefficients = {'first':
                        ((.344, -.208, .064),
                         (.248, -.056, .008),
                         (.176, .048, -.024),
                         (.128, .104, -.032),
                         (.104, .122, -.016)),
                    'middle':
                        ((.064, .152, -.016),
                         (.008, .224, -.032),
                         (-.024, .248, -.024),
                         (-.032, .224, .008),
                         (-.016, .152, .064)),
                    'last':
                        ((-.016, .112, .104),
                         (-.032, .104, .128),
                         (-.024, .048, .176),
                         (.008, -.056, .248),
                         (.064, -.208, .344))}
    if group_index < 0:
        print('Group index cannot be negative')
    elif group_index > last_age_group_index:
        print('Group index cannot be greater than number of groups')
    elif group_index == 0:
        group, group_start_adjustment = 'first', 0
    elif group_index == last_age_group_index:
        group, group_start_adjustment = 'last', -2
    else:
        group, group_start_adjustment = 'middle', -1
    interpolated_estimate = 0.
    for n_age_group in range(3):
        interpolated_estimate += age_group_width * coefficients[group][within_group_index][n_age_group] \
                                 * data[group_index + n_age_group + group_start_adjustment]
    return interpolated_estimate


''' objects '''


class Spring:
    def __init__(self):
        """
        Basic data processing structure that reads the input spreadsheets, processes them and can then be fed to the
        outputs structure for graphing, etc.

        For data structures, dimensions are:
        1. age group
        2. years
        3. gender
        4. cause of death
        """

        # specify spreadsheets to read and read them into single data structure - always put all-causes-combined first
        self.grim_sheets_to_read = ['all-causes-combined', 'all-diseases-of-the-circulatory-system',
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

        self.integer_ages = range(90)
        self.life_tables = {}
        self.cumulative_deaths_by_cause = {}
        self.grim_books_data = {'population': {}, 'deaths': {}}
        self.rates = {}
        self.average_rates_by_year = {}
        self.upper_age_limits_to_cut_at = ['70 to 74', '75 to 79']

        # read population data
        book = open_workbook('grim-' + self.grim_sheets_to_read[0] + '-2017.xlsx')
        (self.grim_books_data['population']['age_groups'], self.grim_books_data['population']['years'],
         self.grim_books_data['population']['genders'], self.grim_books_data['population']['data'], _) \
            = read_grim_sheet(book, 'Populations', title_row_index=14, gender_row_index=12)

        # read death data spreadsheets
        (self.grim_books_data['deaths']['age_groups'], self.grim_books_data['deaths']['years'],
         self.grim_books_data['deaths']['genders'], self.grim_books_data['deaths']['data']) \
            = read_all_grim_sheets(self.grim_sheets_to_read)

        # read in an process the Australian standard 2001 population data
        self.standard_population_data = read_standard_population()
        self.revised_population_data = sum_last_elements_of_dict(sum_dict_over_brackets(self.standard_population_data))
        self.standard_population_props \
            = {key: self.revised_population_data[key] / float(sum(self.revised_population_data.values()))
               for key in self.revised_population_data}

        self.upper_age_limits_to_cut_at.append(self.grim_books_data['deaths']['age_groups'][-2])

        # restrict input array and find relevant years
        self.grim_books_data['deaths']['adjusted_data'] \
            = distribute_missing_across_agegroups(self.grim_books_data['deaths']['data'],
                                                  self.grim_books_data['deaths']['age_groups'])
        self.grim_books_data['population']['adjusted_data'] \
            = restrict_population_to_relevant_years(self.grim_books_data['population']['data'],
                                                    self.grim_books_data['deaths']['years'],
                                                    self.grim_books_data['population']['years'])

        # find death rates from tidied arrays
        self.rates['unadjusted'] \
            = find_rates_from_deaths_and_populations(self.grim_books_data['deaths']['adjusted_data'],
                                                     self.grim_books_data['population']['adjusted_data'],
                                                     len(self.grim_sheets_to_read))

        self.upper_age_limits_to_cut_at.append(self.grim_books_data['deaths']['age_groups'][-2])

        # find average rates across groups
        self.find_average_rates_by_year()

    def find_average_rates_by_year(self):
        """
        Find death rates by year averaged over age groups, but excluding the highest ones.
        """

        g = self.grim_books_data['deaths']['genders'].index('Persons')
        self.average_rates_by_year['adjusted_data'] = {}
        # self.average_rates_by_year['weighted_adjusted_data'] = {}
        for upper_age_limit in self.upper_age_limits_to_cut_at:
            up = self.grim_books_data['deaths']['age_groups'].index(upper_age_limit)
            denominator = numpy.sum(self.grim_books_data['population']['adjusted_data'][:up, :, g], axis=0)
            self.average_rates_by_year['adjusted_data'][upper_age_limit] = {}
            # self.average_rates_by_year['weighted_adjusted_data'][upper_age_limit] = {}
            for c, cause in enumerate(self.grim_sheets_to_read):
                numerator = numpy.sum(self.grim_books_data['deaths']['adjusted_data'][:up, :, g, c], axis=0)
                self.average_rates_by_year['adjusted_data'][upper_age_limit][cause] \
                    = numpy.divide(numerator, denominator)
                # self.average_rates_by_year['weighted_adjusted_data'][upper_age_limit][cause] \
                #     = numpy.zeros(len(denominator))
                # for a, agegroup in enumerate(self.grim_books_data['deaths']['age_groups'][:up]):
                #     self.average_rates_by_year['weighted_adjusted_data'][upper_age_limit][cause] \
                #         = [r + n / d * w for r, n, d, w
                #            in zip(self.average_rates_by_year['weighted_adjusted_data'][upper_age_limit][cause],
                #                   self.grim_books_data['deaths']['adjusted_data'][a, :, g, c],
                #                   self.grim_books_data['population']['adjusted_data'][a, :, g],
                #                   [1. / 19.] * len(self.grim_books_data['deaths']['age_groups'][:up]))]

    def find_life_tables(self, karup_king=True):
        """
        Use the death rates to estimate the remaining proportion left alive and the cumulative deaths by age.

        Args:
            karup_king: Whether to use Karup-King interpolators, rather than rectangular distributions
        Creates:
            self.lifetables: Survivors by age
            self.cumulative_deaths_by_cause: Cumulative deaths by age
        """

        # construct life tables and cumulative death structures for each calendar year
        age_group_lower, age_group_upper \
            = find_agegroup_values_from_strings(self.grim_books_data['deaths']['age_groups'])
        for year in range(self.grim_books_data['deaths']['years'][0], self.grim_books_data['deaths']['years'][-1] + 1):

            # the life table list and the running value to populate it
            survival_total = 1.
            self.life_tables[year] = [1.]

            # the cumulative death structures and the value to populate it, by cause of death
            self.cumulative_deaths_by_cause[year] = {}
            for cause in self.grim_sheets_to_read:
                self.cumulative_deaths_by_cause[year][cause] = [0.]
                cumulative_deaths = 0.

                # looping over each age group
                for age in self.integer_ages:
                    age_group_index = next(x[0] for x in enumerate(age_group_upper) if x[1] >= age)
                    within_group_age = age - age_group_lower[age_group_index]

                    # find rate, either with Karup-King interpolation or without
                    if karup_king:
                        rate_for_age \
                            = karup_king_interpolation(
                                age_group_index, within_group_age, 17,
                                self.rates['unadjusted'][:,
                                self.grim_books_data['deaths']['years'].index(year),
                                self.grim_books_data['deaths']['genders'].index('Persons'),
                                self.grim_sheets_to_read.index(cause)])
                    else:
                        rate_for_age = self.rates['unadjusted'][
                            age_group_index,
                            self.grim_books_data['deaths']['years'].index(year),
                            self.grim_books_data['deaths']['genders'].index('Persons'),
                            self.grim_sheets_to_read.index(cause)]

                    # decrement survival and increment cumulative deaths
                    if cause == 'all-causes-combined':
                        survival_total *= 1. - rate_for_age
                        self.life_tables[year].append(survival_total)
                    else:
                        cumulative_deaths += self.life_tables[year][age] * rate_for_age
                        self.cumulative_deaths_by_cause[year][cause].append(cumulative_deaths)


class Outputs:
    def __init__(self, data_object):
        """
        Outputs module that creates plots based on the data object.

        Args:
            data_object: The data structure with death rates and other information required for plotting
        """

        self.data_object = data_object

    def plot_death_rates_over_time(self, cause='all-causes-combined', x_limits=None, y_limits=(2e-5, .18),
                                   log_scale=False, split_by_gender=True):
        """
        Create graph of total death rates by age groups over time.

        Args:
            cause: String for cause to be plotted
            x_limits: Tuple containing the two elements for the left and right boundary of the x-axis
            y_limits: Tuple containing the two elements for the lower and upper boundary of the y-axis
            log_scale: Whether to plot with a vertical log scale or just linear (if False)
            split_by_gender: Whether to produce three graphs for males, females and both genders
        """

        figure = plt.figure()
        if not x_limits:
            x_limits = (float(min(self.data_object.grim_books_data['deaths']['years'])),
                        float(max(self.data_object.grim_books_data['deaths']['years'])))
        genders_to_plot = self.data_object.grim_books_data['deaths']['genders'] if split_by_gender else ['Persons']

        for gender in genders_to_plot:
            ax = figure.add_axes([0.1, 0.1, 0.6, 0.75])
            gender_string = convert_grim_string(gender) if split_by_gender else ''
            iterations = len(self.data_object.grim_books_data['deaths']['age_groups']) - 1
            colours = [plt.cm.Blues(x) for x in numpy.linspace(0., 1., iterations)]
            year_values = self.data_object.grim_books_data['deaths']['years']
            for i in range(5, iterations):
                rates = self.data_object.rates['unadjusted'][i, :,
                        self.data_object.grim_books_data['deaths']['genders'].index(gender),
                        self.data_object.grim_sheets_to_read.index(cause)]
                label = self.data_object.grim_books_data['deaths']['age_groups'][i]
                if log_scale:
                    ax.semilogy(year_values, rates, label=label, color=colours[i])
                else:
                    ax.plot(year_values, rates, label=label, color=colours[i])
            handles, labels = ax.get_legend_handles_labels()
            ax.legend(handles, labels, bbox_to_anchor=(1.05, 1), loc=2, borderaxespad=0., frameon=False,
                      prop={'size': 7})
            ax.set_title(gender_string)
            ax.set_ylim(y_limits)
            ax.set_xlim(x_limits)
        plt.setp(ax.get_xticklabels(), fontsize=10)
        plt.setp(ax.get_yticklabels(), fontsize=10)
        figure.savefig('mortality_figure_' + gender.lower())

    def plot_deaths_by_cause(self):
        """
        Deaths by cause with limitation by age group and with age groups unlimited.
        """

        for upper_age_limit in self.data_object.upper_age_limits_to_cut_at:
            upper_age_limit_string = '' \
                if upper_age_limit == self.data_object.grim_books_data['deaths']['age_groups'][-2] \
                else ', under ' + upper_age_limit[:2] + 's'

            figure = plt.figure()
            ax = figure.add_axes([0.1, 0.1, 0.6, 0.75])
            for cause in self.data_object.grim_sheets_to_read:
                ax.plot(self.data_object.grim_books_data['deaths']['years'],
                        self.data_object.average_rates_by_year['adjusted_data'][upper_age_limit][cause],
                        label=convert_grim_string(cause))
            handles, labels = ax.get_legend_handles_labels()
            ax.legend(handles, labels, bbox_to_anchor=(1.05, 1), loc=2, borderaxespad=0., frameon=False,
                      prop={'size': 7})
            ax.set_title('Death rates by cause' + upper_age_limit_string)
            ax.set_ylim((0., 9e-3))
            ax.set_xlim((1964., 2014.))
            ax.set_xlabel('Year', fontsize=10)
            ax.set_ylabel('Rate per capita per year', fontsize=10)
            plt.setp(ax.get_xticklabels(), fontsize=10)
            plt.setp(ax.get_yticklabels(), fontsize=10)
            figure.savefig('mortality_figure_cause' + upper_age_limit_string)

    def plot_cumulative_survival(self):
        """
        Plot cumulative survival graphs by year and age.
        """

        figure = plt.figure()
        n_plots, rows, columns, base_font_size, year_spacing, last_year = 3, 2, 2, 8, 25, 2014
        plt.style.use('ggplot')
        for n_plot in range(n_plots):
            year = last_year + n_plot * year_spacing - (n_plots - 1) * year_spacing
            ax = figure.add_subplot(rows, columns, n_plot + 1)
            stacked_data = {'base': numpy.zeros(len(self.data_object.life_tables[year])),
                            'survival': self.data_object.life_tables[year],
                            'cumulative other deaths': numpy.ones(len(self.data_object.life_tables[year]))}
            ordered_list_of_stacks = ['base', 'survival']
            new_data = self.data_object.life_tables[year]
            for cause in self.data_object.cumulative_deaths_by_cause[year]:
                if cause != 'all-causes-combined':
                    new_data = [i + j for i, j in zip(new_data, self.data_object.cumulative_deaths_by_cause[year][cause])]
                    stacked_data[cause] = new_data
                    ordered_list_of_stacks.append(cause)
            ordered_list_of_stacks.append('cumulative other deaths')
            for i in range(1, len(ordered_list_of_stacks)):
                ax.fill_between(self.data_object.integer_ages,
                                stacked_data[ordered_list_of_stacks[i - 1]][:-1],
                                stacked_data[ordered_list_of_stacks[i]][:-1],
                                color=list(plt.rcParams['axes.prop_cycle'])[i - 1]['color'],
                                label=find_string_from_dict(ordered_list_of_stacks[i]))
            handles, labels = ax.get_legend_handles_labels()
            if n_plot >= columns:
                ax.set_xlabel('Age', fontsize=base_font_size)
            if n_plot % columns == 0:
                ax.set_ylabel('Proportion', fontsize=base_font_size)
            if n_plot == n_plots - 1:
                ax.legend(handles, labels, bbox_to_anchor=(1.13, .8), loc=2, frameon=False, prop={'size': 9})
            plt.setp(ax.get_xticklabels(), fontsize=base_font_size - 2)
            plt.setp(ax.get_yticklabels(), fontsize=base_font_size - 2)
            ax.set_title(year, fontsize=base_font_size + 2)
            ax.set_xlim((50., 89.))
            ax.set_ylim((0., 1.))
        plt.tight_layout()
        figure.savefig('lifetable')

