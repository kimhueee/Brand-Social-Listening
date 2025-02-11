import pandas as pd

def hyperlink(df, original_df, topORcom):
    # Take the top 10 most frequent SiteName/Author:
    df = pd.DataFrame(df).sort_values(by='Id', ascending=False).head(10)
    # Necessary variables declaration:
    out_put = pd.DataFrame()
    if topORcom == 'top':
        topic_or_comment = 'UrlTopic'
    elif topORcom == 'com':
        topic_or_comment = 'UrlComment'
    else:
        print('Wrong input of topORcom')
    # Create the output Dataframe:
    for each_record in list(df.index):
        link_df = original_df[original_df[df.index.name] == each_record].pivot_table(index=topic_or_comment, values='Id', aggfunc='count').sort_values(by='Id', ascending=False)
        out_put.loc[each_record, 'Link'] = ('=' + 'HYPERLINK(' + '"' + str(list(link_df.index)[0]) + '",' + '"' + str(each_record) + '")')
    out_put.index.name = df.index.name

    out_put = pd.merge(out_put, df, on=df.index.name, how='left').set_index('Link')

    return out_put

