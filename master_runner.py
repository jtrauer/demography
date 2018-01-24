
from grim_reader import Spring, Outputs

data_object = Spring()
data_object.find_life_tables()

outputs_object = Outputs(data_object)
outputs_object.plot_death_rates_over_time(x_limits=(1960., 2014.))
outputs_object.plot_deaths_by_cause()
# outputs_object.plot_cumulative_survival()
