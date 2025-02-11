import pandas as pd
import vaex as vx
from datemerge import datemerge
from hyperlink import hyperlink
from to_excel_file import to_excel_file

# file_name = r"C:\Users\Admin\Desktop\SPF1723Sample.xlsx"

file_dir1 = r"C:\Users\Kompa-KTelt056\Desktop\spfdec.xlsx"
file_dir2 = r"C:\Users\Kompa-KTelt056\Desktop\spfcomdec.xlsx"
# file_dir3 = r"C:\Users\Admin\Desktop\spfcommar2.xlsx"


df1 = vx.from_pandas(pd.read_excel(file_dir1, sheet_name='Data'))
# df2 = vx.from_pandas(pd.read_excel(file_dir1, sheet_name='mg2'))
# df3 = vx.from_pandas(pd.read_excel(file_dir1, sheet_name='mg3'))
# df4 = vx.from_pandas(pd.read_excel(file_dir1, sheet_name='mg4'))
# df5 = vx.from_pandas(pd.read_excel(file_dir1, sheet_name='mg5'))

df7 = vx.from_pandas(pd.read_excel(file_dir2, sheet_name='Data'))
# df8 = vx.from_pandas(pd.read_excel(file_dir2, sheet_name='25to30'))

# df9 = vx.from_pandas(pd.read_excel(file_dir3, sheet_name='01to19'))
# df6 = vx.from_pandas(pd.read_excel(file_dir3, sheet_name='Data'))



df = vx.concat([df1, df7])

df_excluded_minigame = df[(df.Minigame != 'Minigame') & (df.Minigame != 'minigame')]

label_classify_df = pd.read_excel(r"C:\Users\Kompa-KTelt056\Desktop\SPFlabels.xlsx", sheet_name='Label classify')

sentiment_list = ['Positive', 'Neutral', 'Negative']
channel_list = ['Facebook', 'Owned channel', 'Forum', 'News', 'Youtube', 'Tiktok', 'Ecommere']
target_mother_label_list = ['Service', 'Fanpage Activity', 'Campaign', 'Rider', 'Merchant', 'App Experience']
topic_list = ['ShopeeFood', 'GrabFood', 'BeFood']
slide_n_df_dict = {}
slide_counter = 4

# ============================= SLIDE 4 (TOTAL BUZZ):
buzz_df = df.groupby(by='Topic').agg({'Id': 'count'}).to_pandas_df().set_index('Topic').reindex(topic_list)
buzz_df.loc['Total', 'Id'] = buzz_df['Id'].sum().sum()
slide_n_df_dict['Slide 4'] = buzz_df
slide_counter += 1

# ============================= SLIDE 5 (BUZZ EXCLUDE MINIGAME):
buzz_excluded_minigame_df = df_excluded_minigame.groupby(by='Topic').agg({'Id': 'count'}).to_pandas_df().set_index('Topic').reindex(topic_list)
buzz_excluded_minigame_df.loc['Total', 'Id'] = buzz_excluded_minigame_df['Id'].sum().sum()
slide_n_df_dict['Slide 5'] = buzz_excluded_minigame_df
slide_counter += 1

# ============================= SLIDE 6 (BUZZ EXCLUDE MINIGAME AND REQUESTED SITES):
df_excluded_minigame_n_sites = df_excluded_minigame[(df_excluded_minigame.SiteName != 'Hà Nội: Ăn gì? Ở đâu?') & (df_excluded_minigame.SiteName != 'Sài Gòn: Ăn gì? Ở đâu?') & (df_excluded_minigame.SiteName != 'Đà Nẵng: Ăn gì? Ở đâu?')]
buzz_excluded_minigame_n_sites_df = df_excluded_minigame_n_sites.groupby(by='Topic').agg({'Id': 'count'}).to_pandas_df().set_index('Topic').reindex(topic_list)
buzz_excluded_minigame_n_sites_df.loc['Total', 'Id'] = buzz_excluded_minigame_n_sites_df['Id'].sum().sum()
slide_n_df_dict['Slide 6'] = buzz_excluded_minigame_n_sites_df
slide_counter += 1

# ============================= SLIDE 7, 8 (BUZZ/ EXCLUDE MINIGAME BY CHANNELS):
for wk_df in [df, df_excluded_minigame]:

    # Buzz by channel distribution:
    buffer_df_list = []
    buzz_by_channel = wk_df.groupby(by='Channel').agg({'Id': 'count'}).to_pandas_df().set_index('Channel').reindex(channel_list)
    buzz_by_channel.loc['Total', :] = buzz_by_channel['Id'].sum().sum()
    buzz_by_channel.iloc[:-1, 0] = buzz_by_channel.iloc[:-1, 0].apply(lambda x: round(x/buzz_by_channel.loc['Total', 'Id'], 4))
    buffer_df_list.append(buzz_by_channel)

    # Buzz by channel distribution:
    horizontal_channel_df = wk_df.groupby(by='Channel').agg({'Id': 'count'}).to_pandas_df().pivot_table(columns='Channel', values='Id', aggfunc='sum').reindex(channel_list, axis=1)
    buffer_df_list.append(horizontal_channel_df)

    # Buzz by channel distribution:
    channel_distribution_by_topic_df = wk_df.groupby(by=['Topic', 'Channel']).agg({'Id': 'count'}).to_pandas_df().pivot_table(index='Topic', columns='Channel', values='Id', aggfunc='sum').reindex(topic_list).reindex(channel_list, axis=1)
    channel_distribution_by_topic_df = channel_distribution_by_topic_df.apply(lambda x: round(x/channel_distribution_by_topic_df.sum(axis=1), 4))
    buffer_df_list.append(channel_distribution_by_topic_df)

    if wk_df is df_excluded_minigame:
        buffer_df_list.append(wk_df.groupby(by=['Topic', 'Channel']).agg({'Id': 'count'}).to_pandas_df().pivot_table(index='Topic', columns='Channel', values='Id', aggfunc='sum').reindex(topic_list).reindex(channel_list, axis=1))

    slide_n_df_dict['Slide ' + str(slide_counter)] = buffer_df_list
    slide_counter += 1

# ============================= SLIDE 9 (BUZZ BY SENTIMENT):
    Slide9_df_list = []

    sent_df = df.groupby(by='Sentiment').agg({'Id': 'count'}).to_pandas_df().set_index('Sentiment').reindex(sentiment_list)
    sent_df = sent_df.apply(lambda x: round(x/sent_df.sum().sum(), 4))

    Slide9_df_list.append(sent_df)
    slide_n_df_dict['Slide ' + str(slide_counter)] = sent_df

# ============================= SLIDE 10 AND 13 TO 35 (BUZZ SUMMARY & CONVERSATION ANALYSIS BY BRAND):
slide_counter += 1
Slide10_df_list = []
slide13to35_counter = 13

for brand in topic_list:
    Slide13to35_df_list = []

    # FILTER DATA FRAME BY BRAND:
    brand_working_df = df[df['Topic'] == brand]
    constant = int(brand_working_df['Id'].count())/int(brand_working_df['Sentiment'].count())
    all_labels_df = pd.DataFrame(columns=sentiment_list)

    # BUZZ BY CHANNEL BY BRAND (SLIDE 13):
    buzz_by_channel_by_brand_df = brand_working_df.groupby(by='Channel').agg({'Id': 'count'}).to_pandas_df().pivot_table(columns='Channel', values='Id', aggfunc='sum').reindex(channel_list, axis=1)
    buzz_by_channel_by_brand_df.index.name = brand
    Slide13to35_df_list.append(buzz_by_channel_by_brand_df)

    # TRENDLINE (SLIDE 13):
    Slide13to35_df_list.append(datemerge(brand_working_df.groupby(by='PublishedDate').agg({'Id': 'count'}).to_pandas_df().pivot_table(
        index='PublishedDate', values='Id', aggfunc='sum'), '2024-12-01', '2024-12-31'))

    # ADD TO DICTIONARY OF SLIDE 13 TO 35:
    slide_n_df_dict['Slide ' + str(slide13to35_counter)] = Slide13to35_df_list
    slide13to35_counter += 1
    Slide13to35_df_list = []

    # BUZZ BY SENTIMENT BY BRAND:
    sent_by_brand_df = brand_working_df.groupby(by='Sentiment').agg({'Id': 'count'}).to_pandas_df().set_index('Sentiment').reindex(sentiment_list)
    sent_by_brand_df.loc['Total', :] = sent_by_brand_df.sum().sum()
    sent_by_brand_df.iloc[:-1, 0] = sent_by_brand_df.iloc[:-1, 0].apply(lambda x: round(x / sent_by_brand_df.loc['Total', 'Id'], 4))
    sent_by_brand_df.index.name = brand

    Slide10_df_list.append(sent_by_brand_df)
    Slide13to35_df_list.append(sent_by_brand_df)

    # LABELS DISTRIBUTION BY SENTIMENT:
        # Consolidate labels by sentiment from even columns:
    for col_name in ['Service', 'Campaign', 'Labels2', 'Labels4', 'Labels6', 'Labels8', 'Labels10']:
        all_labels_df = pd.concat([all_labels_df, brand_working_df.groupby([col_name, 'Sentiment']).agg({'Id': 'count'}).to_pandas_df().pivot_table(index=col_name, columns='Sentiment', values='Id', aggfunc='sum').reindex(sentiment_list, axis=1).fillna(0)])

        # Multiply the whole dataframe by the constant :
    all_labels_df = pd.DataFrame(all_labels_df).apply(lambda x: x.astype(int) * constant).reset_index().rename({'index': 'SonLabels'}, axis=1).merge(label_classify_df, on='SonLabels', how='left')
    all_labels_df = pd.DataFrame(all_labels_df.groupby(['MotherLabels', 'SonLabels']).sum()).reindex(sentiment_list, axis=1)
    all_labels_df['Total'] = all_labels_df.sum(axis=1)
    all_labels_df = all_labels_df.reset_index().merge(pd.DataFrame(all_labels_df.groupby(['MotherLabels'])['Total'].sum()).reset_index().rename({'Total': 'TotalByMoLabels'}, axis=1), on='MotherLabels', how='left').sort_values(by=['TotalByMoLabels', 'Total'], ascending=[True, True]).drop('TotalByMoLabels', axis=1)

        # Groupby Mother_Labels:
    mother_labels_df = all_labels_df.groupby('MotherLabels').sum().sort_values(by='Total', ascending=True).apply(lambda x: round(x/int(brand_working_df['Id'].count()), 5))
    mother_labels_df.index.name = brand

    Slide10_df_list.append(mother_labels_df)
    Slide13to35_df_list.append(mother_labels_df)

        # Groupby Target Mother Labels:
    for mother_label in target_mother_label_list:
        target_mother_label_df = all_labels_df[all_labels_df['MotherLabels'] == mother_label]
        sonlabels_by_sent_df = target_mother_label_df.groupby('SonLabels').sum().reindex(sentiment_list, axis=1)
        sonlabels_by_sent_df = sonlabels_by_sent_df.apply(lambda x: round(x / sonlabels_by_sent_df.sum().sum(), 3))
        sonlabels_by_sent_df['Total'] = sonlabels_by_sent_df.sum(axis=1)
        sonlabels_by_sent_df.index.name = mother_label

        Slide13to35_df_list.append(sonlabels_by_sent_df)
    slide_n_df_dict['Slide ' + str(slide13to35_counter)] = Slide13to35_df_list
    slide13to35_counter += 1
    Slide13to35_df_list = []

    # TOP SOURCES:
    for sent in ['Positive', 'Negative']:
            # Filter the brand Data frame by positive/ negative buzz:
        brand_by_sent_working_df = brand_working_df[brand_working_df['Sentiment'] == sent]
            # Define the top 10 sources:
        top10_sources_df = brand_by_sent_working_df.groupby(by='SiteName').agg({'Id': 'count'}).sort(by='Id', ascending=False).to_pandas_df().head(10).set_index('SiteName')
        top10_sources_df['Link'] = None
        top10_sources_df['Channel'] = None
            # For each site, define the most frequent 'UrlTopic':
        for site in list(top10_sources_df.index):
            brand_by_sent_site_working_df = brand_by_sent_working_df[brand_by_sent_working_df['SiteName'] == site]
            # Define the most frequent UrlTopic:
            top10_sources_df.loc[site, 'Link'] = list(brand_by_sent_site_working_df.groupby(by='UrlTopic').agg({'Id': 'count'}).sort(by='Id', ascending=False).head(1)['UrlTopic'].unique())[0]
            # Define the Channel:
            top10_sources_df.loc[site, 'Channel'] = list(brand_by_sent_site_working_df['Channel'].unique())[0]

        top10_sources_df = top10_sources_df.reset_index()
        top10_sources_df['Id'] = top10_sources_df['Id'].apply(lambda x: round(x/top10_sources_df['Id'].sum().sum(), 5))

        for i in range(len(top10_sources_df.axes[0])):
            top10_sources_df.iloc[i, 0] = ('=' + 'HYPERLINK(' + '"' + str(top10_sources_df.iloc[i, 2]) + '",' + '"' + str(top10_sources_df.iloc[i, 0]) + '")')

        top10_sources_df = top10_sources_df[['SiteName', 'Id', 'Channel']].set_index(['SiteName'])
        top10_sources_df.index.name = str(sent)
        Slide13to35_df_list.append(top10_sources_df)

    # ADD TO DICTIONARY OF SLIDE 13 TO 35 (ADD RESPECTIVE SLIDES OF EACH BRAND):
    slide_n_df_dict['Slide ' + str(slide13to35_counter)] = Slide13to35_df_list
    slide13to35_counter += 3

slide_n_df_dict['Slide ' + str(slide_counter)] = Slide10_df_list
slide_counter += 1

# ============================= SLIDE 11 (NSR):

Slide11_df_list = []
buzz_nsr_author_df = df.groupby(['Topic', 'Sentiment']).agg({'Id': 'count'}).to_pandas_df()

    # THE CHART:
nsr_chart_df = buzz_nsr_author_df.pivot_table(index='Topic', columns='Sentiment', values='Id', aggfunc='sum').fillna(0)
nsr_chart_df['NSR'] = round((nsr_chart_df['Positive'] - nsr_chart_df['Negative']) / (nsr_chart_df['Positive'] + nsr_chart_df['Negative']), 3)
nsr_chart_df = nsr_chart_df.drop(sentiment_list, axis=1)

buzz_vertical_df = df.groupby('Topic').agg({'Id': 'count'}).to_pandas_df().pivot_table(index='Topic', values='Id', aggfunc='sum')

unique_voice_vertical_df = df.groupby('Topic').agg({'AuthorId': 'nunique'}).to_pandas_df().pivot_table(index='Topic', values='AuthorId', aggfunc='sum')

Slide11_df_list.append(pd.concat([buzz_vertical_df, nsr_chart_df, unique_voice_vertical_df], axis=1).reindex(topic_list))

    # THE TABLE:
nsr_table_df = buzz_nsr_author_df.pivot_table(index='Sentiment', columns='Topic', values='Id', aggfunc='sum')
nsr_table_df.loc['NSR', :] = round((nsr_table_df.loc['Positive', :] - nsr_table_df.loc['Negative', :]) / (nsr_table_df.loc['Positive', :] + nsr_table_df.loc['Negative', :]), 3)
nsr_table_df = nsr_table_df.drop(sentiment_list).fillna(0)

buzz_horizontal_df = df.groupby('Topic').agg({'Id': 'count'}).to_pandas_df().pivot_table(columns='Topic', values='Id', aggfunc='sum')

unique_voice_horizontal_df = df.groupby('Topic').agg({'AuthorId': 'nunique'}).to_pandas_df().pivot_table(columns='Topic', values='AuthorId', aggfunc='sum')

Slide11_df_list.append(pd.concat([buzz_horizontal_df, nsr_table_df, unique_voice_horizontal_df], axis=0).reindex(topic_list, axis=1))

    # Add Slide 11:
slide_n_df_dict['Slide ' + str(slide_counter)] = Slide11_df_list
slide_counter += 2


# ============================= SORT DICTIONARY AND ADD TO EXCEL FILE:

    # SORT BY KEYS:
myKeys = list(slide_n_df_dict.keys())

sort_myKeys = sorted([int(words[words.find(' '):]) for words in myKeys], reverse=False)

slide_n_df_dict = {'Slide ' + str(i): slide_n_df_dict['Slide ' + str(i)] for i in sort_myKeys}

    # ADD TO EXCEL FILE:
to_excel_file('SPFdecOutPut', slide_n_df_dict)
