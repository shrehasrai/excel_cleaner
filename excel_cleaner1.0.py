
import pandas as pd
from configparser import ConfigParser






ipl_data = {'session_title': ['Red', 'Red', 'Red', 'Blue','Blue', 'Blue'],
   'author_affiliation': ['US', 'India', '', 'Mumbai', 'Pasika; Bakes', ''],
   'authors': ['Shivam','Sunny','All Participants','Sejal', 'Neera; buccha', 'All Participants'],
    'Gender': ['Boy', 'Girl','Boy', 'Girl', 'Girl', 'Girl']}



df = pd.DataFrame(ipl_data)

print(df)
print('_______________________')
import sys
#sys.exit()


#Read config.ini file
config_object = ConfigParser()
config_object.read("config.ini")

#Get the password
details = config_object["details"]
excel_name = details["target"]
# allow_unique_author = details["allow_unique_author"]
manual_date_format = details["manual_date_format"]
author_affiliation_special_end_removal = details["author_affiliation_special_end_removal"]
sheet_name = details['sheet_name']
group_by = details['group_by']
group_by_author_replace = details['group_by_author_replace']


def convert_df(name):

    df = pd.read_excel(f'{name}.xlsx',engine='openpyxl')
    df.reset_index(drop=True, inplace=True)
    return df


df = convert_df(excel_name)

def chk_aff_already_added(to_find_aut, to_find_aff, aut_list, aff_list)-> bool:
    '''
    True :- aff and aut already exist
    False :- aff not added
    '''

    for aut, aff in zip(aut_list, aff_list):
        if aut == to_find_aut and aff == to_find_aff:
            return True

    return False


# filling all authors
if group_by_author_replace.lower() != 'false':

#     group_by = "session_title"#enter the name of column to group by
#     group_by_author_replace = "All Participants"#eg all authors


    #if not (group_by in df):
    #    print(f'{group_by} is not available in excel file')
    #    print(f'try using one from {list(df.columns)}')

    group_by_lst = group_by.split('=>')

    print(df.columns)    
    grp = df.groupby(group_by_lst)


    for g in grp:
        if str(g[0]) and str(g[0]).lower() != 'nan':#avoiding empty values and nan values
            g_df = g[1].copy()
            all_authors = g_df['authors'].astype(str).to_list()
            all_aff = g_df['author_affiliation'].astype(str).to_list()


            final_aut = []
            final_aff = []

            for aut,aff in zip(all_authors, all_aff):
                for k,v in zip(aut.split('; '), aff.split('; ')):
                    if k != 'nan' and k != group_by_author_replace:#if aut is valid
                        if v == 'nan':
                            if k not in final_aut:
                                final_aut.append(k)
                                final_aff.append('')
                        else:
                            if not chk_aff_already_added(k, v, final_aut, final_aff):
                                final_aut.append(k)
                                final_aff.append(v)
    #         uncoment to change
#             print(all_authors)
#             print(all_aff)
#             print('after processing---------------')
#             print('; '.join(list(_dict.keys())))
#             print('; '.join(list(_dict.values())))
#             print('\n____________________')

            temp_df = g[1][g[1]['authors'] == group_by_author_replace]
            
            df.loc[temp_df.index, ['author_affiliation']] = '; '.join(final_aff)

            df.loc[temp_df.index, ['authors']] = '; '.join(final_aut)


df.to_csv("target_colored.csv", index=False)
print(df)
    
