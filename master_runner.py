
# import main objects
from grim_reader import *
import itertools

# run data analysis
data_object = Spring()
# data_object.find_life_tables()

# run plotting through outputs module
outputs_object = Outputs(data_object)
# outputs_object.plot_rates_by_age_group_over_time(
#     cause='liver-disease', x_limits=(1980., 2014.), split_by_gender=False)
# outputs_object.plot_journal_figure_1()
# outputs_object.plot_cumulative_survival()
# outputs_object.plot_deaths_by_cause()

# outputs for the aspree paper

aspree_data = {
    70: {'Persons': 9668, 'Females': 5173},
    75: {'Persons': 4432, 'Females': 2515},
    80: {'Persons': 1963, 'Females': 1125},
    85: {'Persons': 640, 'Females': 367}}
aspree_total = 0
genders = ['Males', 'Females']
for age in range(70, 90, 5):
    aspree_total += aspree_data[age]['Persons']
    aspree_data[age]['Males'] = aspree_data[age]['Persons'] - aspree_data[age]['Females']
    for gender in genders:
        aspree_data[age]['Proportion ' + gender] = aspree_data[age][gender] / aspree_total
aspree_weights = {}
for age, gender in itertools.product(range(70, 90, 5), genders):
    aspree_weights[str(age) + ' ' + gender] = float(aspree_data[age][gender]) / float(aspree_total)
for year in range(2014, 2017):
    weighted_rate = 0.
    for gender, age in itertools.product(genders, range(70, 90, 5)):
        cancer_deaths = outputs_object.get_rate(
                convert_integer_age_to_string(age), year, gender, 'all-neoplasms', 'raw_deaths')
        population = outputs_object.get_rate(convert_integer_age_to_string(age), year, gender, '', 'population')
        rate = cancer_deaths / population
        contribution = rate * aspree_weights[str(age) + ' ' + gender]
        weighted_rate += contribution
    print('\nWeighted rate in {} is:'.format(year))
    print(weighted_rate * 1e3)


