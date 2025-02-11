import pandas as pd
from datetime import datetime, timedelta

def datemerge(ref_df, start_date, end_date):

        # Define start/ end date of the expected period:
    start_date = datetime.strptime(start_date, '%Y-%m-%d')
    end_date = datetime.strptime(end_date, '%Y-%m-%d')

        # Create the Dataframe with index is the list of date from expected period:
    df = pd.DataFrame(index=[start_date + timedelta(days=day) for day in range((end_date - start_date).days + 1)])
    df.index.name = 'PublishedDate'

    output = pd.merge(df, ref_df, on='PublishedDate', how='left').fillna(value=0)

    return output





