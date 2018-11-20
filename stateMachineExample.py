from transitions import Machine

class ReadCode(object):
    pass

code = ReadCode()


# The states
states=[
    'init',
    'search_code',
    'write_sensor_list',
    'write_coeff_table',
    'write_cat_title_list',
    'write_cat_table',
    'write_smv_title_list',
    'write_SMV_table']


#     trigger                source                dest
transitions = [
    ['coeff_table_found',   'search_code',        'write_sensor_list'],
    ['end_of_coeff_table',  'write_coeff_table',  'search_code'],
    ['cat_table_found',     'search_code',        'write_cat_title_list'],
    ['end_of_cat_table',    'write_cat_table',    'search_code'],
    ['smv_table_found',     'search_code',        'write_smv_title_list'],
    ['end_of_smv_table',    'write_smv_table',    'search_code']
]

# Initialize
machine = Machine(code, states, transitions=transitions, initial='search_code')

#print(code.state)
#code.coeff_table_found()
#print(code.state)



