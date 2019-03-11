
# coding: utf-8

# In[1]:


import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import argparse


# In[ ]:


# argument parser in event command line args are passed
def ArgParser():
    
    parser = argparse.ArgumentParser(description="sequence object passed on the command line")
    
    parser.add_argument("-f",
                        "--file",
                        required=True,
                        help="The file or filepath of the workbook to apply fuzzy matching on.  File should be limited by utterance date and serial num",
                        )
    return parser


# In[ ]:


def fuzzy_match(refs, hyps):
    
    # Goal: performs fuzzy matching based on Levenshtein distance
    
    matches = {}
    
    for index, row in hyps.iterrows():
        try:
            match = process.extract(row['hypothesis'], refs['ref'])
            matches[index] = [row['haystack_handle'], match[0][0]]
        except:
            continue
            
    print('Found {} potential matches.'.format(len(matches)))
    
    df_col = ['haystack_handle2', 'ref']
    df = pd.DataFrame.from_dict(matches, orient='index', columns=df_col)
    df = df.merge(refs, left_on='ref', right_on='ref')
    
    merged = hyps.merge(df, how='left', left_on='haystack_handle', right_on='haystack_handle2')
    merged = merged.drop_duplicates(subset='haystack_handle', keep='first')
    
    return merged


# In[ ]:


if __name__ == "__main__":
    args = ArgParser().parse_args()
    
    file = args.file
    
    print("Opening file. Parsing...")
    
    xlsx = pd.ExcelFile(file)

    # extract DFs from workbook
    hyps = pd.read_excel(xlsx, 'Sheet1')
    refs = pd.read_excel(xlsx, 'Sheet2')
    
    merged = fuzzy_match(refs, hyps)
    
    print('Matched. Writing to file...')
    
    merged.to_excel('oakland_audio_data_'+file)
    
    print("All Done!")
    

