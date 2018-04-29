# created by Vedant Jha on 18/04/2018
# under the mentorship of Dr. Akshit Singh
# Task to perform sentiment analysis on the supply chain incidents present in the sample.xlsx file,
# the input file is assumed "preprocessed.xls" file which was created from the part I
# This is the part-2 of the task.
# This program is divided into the 3 parts:
#    1. Calcualtion of the Aggregate Sentiment Score
#    2. Hierarchical Clustering and Making Dendrograms
#    3. Geo-Graphic visualization

# first importing all the required packages to be used
import folium  # plotting the coordinates on the map
import matplotlib.pyplot as plt  # For plotting the graphs and charts
import numpy as np  # For number computing
import pandas as pd  # To handle data
import tweepy  # To use twitter api to get the twitter user data like the location, profile information
import xlwt  # For creating new spread-sheets and saving the data into the sheets
from geopy import geocoders  # To get the latitude and longitude of the place
from matplotlib import style  # styling the maps
from numpy import array  # modifying lists to array
from scipy.cluster.hierarchy import cophenet  # used in hierarchical clustering
from scipy.cluster.hierarchy import dendrogram, linkage  # used in making dendrograms
from scipy.spatial.distance import pdist  # calculating the correlative distance
from sklearn.cluster import MeanShift

# consumer key, consumer secret, access token, access secret. , for using the tweepy api ,
# to get the consumer key goto apps.twitter.com

ckey = "RLQOMc5vwvw13oIBDuwh34pkk"
csecret = "ygtSjFrYbsj0R416Lt75AM0RcFcYNxzyM0fDvu8zIw9kvOkdjH"
atoken = "767929998426660864-pRzl6dIb9swgsVGILBy3dt9G8yeKreT"
asecret = "KepfirsKGEStjDmpcabXyD4OYvsCEEttHiWEM8wsLai9K"


# API's setup:
def twitter_setup():
    """
    Utility function to setup the Twitter's API
    with our access keys provided.
    """
    # Authentication and access using keys:
    auth = tweepy.OAuthHandler(ckey, csecret)
    auth.set_access_token(atoken, asecret)

    # Return API with authentication:
    api = tweepy.API(auth)
    return api


style.use("ggplot")

list_sheetnames = pd.ExcelFile("preprocessed.xls").sheet_names
total_num_sheets = len(list_sheetnames)


def get7charsfromlist(list_):
    result = []
    for single_string in list_:
        result.append(single_string[:7] + "...")
    return result


def get_polarity_subjectivity_list(polarity_, subjectivity_):
    result = []
    for polarity_item, subjectivity_item in zip(polarity_, subjectivity_):
        temp = [0, 0]
        temp[0] = polarity_item
        temp[1] = subjectivity_item
        result.append(temp)
    return result


style_blue = xlwt.easyxf('font: name Times New Roman, color-index blue, bold on', num_format_str='#, ##0.00')
wb = xlwt.Workbook()  # workbook for saving the aggregate sentiment values
wb2 = xlwt.Workbook()  # workbook for saving the hierarchical clustering data
wb3 = xlwt.Workbook()  # workbook for saving the coordinates and places of the twitter users


def getuser(user_date):
    user = ""
    if user_date == -1:
        return ""
    b = False
    for c in user_date:
        if b and str(c) is not " " and str(c) is not "\xa0":
            user = str(user) + str(c)
        elif str(c) is not " " and str(c) is not "\xa0":
            if str(c) is "@":
                b = True
        else:
            b = False
    return user


def get_sentiment_probablity(list_):
    zero = 0
    pos = 0
    neg = 0
    for item in list_:
        if item > 0:
            pos = pos + 1
        elif item < 0:
            neg = neg + 1
        else:
            zero = zero + 1

    return zero / len(list_), pos / (len(list_)), neg / len(list_)


def get_aggregate_sentiment_score(polarity_, subjectivity_):
    val = 0.0
    for polarity_item, subjectivity_item in zip(polarity_, subjectivity_):
        if round(polarity_item, 2) != 0.00:
            val = val + polarity_item * (1 - subjectivity_item)

    return round(val, 3)


def fancy_dendrogram(*args, **kwargs):
    max_d = kwargs.pop('max_d', None)
    if max_d and 'color_threshold' not in kwargs:
        kwargs['color_threshold'] = max_d
    annotate_above = kwargs.pop('annotate_above', 0)

    ddata = dendrogram(*args, **kwargs)

    if not kwargs.get('no_plot', False):
        for i, d, c in zip(ddata['icoord'], ddata['dcoord'], ddata['color_list']):
            x = 0.5 * sum(i[1:3])
            y = d[1]
            if y > annotate_above:
                plt.plot(x, y, 'o', c=c)
                plt.annotate("%.3g" % y, (x, y), xytext=(0, -5),
                             textcoords='offset points',
                             va='top', ha='center')
        if max_d:
            plt.axhline(y=max_d, c='k')
    return ddata


# variables as the list containing the polarity probability and aggregate sentiment scores
#  for each of the supply chain incidents
list_positive = []  # list containing all the positive sentiment to be used in the calculation in the aggregate sentiment score
list_negative = []  # list containing all the negative sentiment values to be used in the calculation in the aggregate sentiment score
list_neutral = []
list_aggregate_sentiment_score = []

# aggregate sentiment score is calcualted as the summation(polarity*(1-subjectiivty))

# variables for calculating the total values for all the 11 supply chain incidents
total_tweets = 0
total_positive = 0
total_negative = 0
total_neutral = 0
total_sentiment_score = 0

ws = wb.add_sheet("score")
ws.write(0, 0, "S. No. ", style_blue)
ws.write(0, 1, "Supply Chain Incidents", style_blue)
ws.write(0, 2, "Positive Sentiment Probablity", style_blue)
ws.write(0, 3, "Negative Sentiment Probability", style_blue)
ws.write(0, 4, "Neutral Sentiment Probablity", style_blue)
ws.write(0, 5, "Aggt. Sentiment Score", style_blue)

ws2 = wb2.add_sheet("clustering")
ws2.write(0, 0, "S. No. ", style_blue)
ws2.write(0, 1, "Supply Chain Incidents", style_blue)
ws2.write(0, 2, "Polarity(X)", style_blue)
ws2.write(0, 3, "Subjectivity(Y)", style_blue)
ws2.write(0, 4, "Number of Clusters", style_blue)
ws2.write(0, 5, "Cluster Points", style_blue)
ws2.write(0, 6, "Dendrogram Corelation Value ", style_blue)

for i in range(total_num_sheets):
    sheet_read = pd.read_excel("preprocessed.xls", list_sheetnames[i])
    df = pd.DataFrame(sheet_read)
    # getting all the values from the preprocessed.xls file
    polarity = sheet_read["Polarity"].values.tolist()
    subjectivity = sheet_read["Subjectivity"].values.tolist()
    twitter_user_list = sheet_read["User & Date"].values.tolist()
    s_neutral, s_positive, s_negative = get_sentiment_probablity(polarity)
    aggregate_sentiment_score = get_aggregate_sentiment_score(polarity, subjectivity)
    # now writing the probability and aggregate_score in the aggregate_sentiment_score.xls file with the sheet names
    ws.write(i + 1, 0, i + 1)
    ws.write(i + 1, 1, list_sheetnames[i])
    ws.write(i + 1, 2, s_positive)
    ws.write(i + 1, 3, s_negative)
    ws.write(i + 1, 4, s_neutral)
    ws.write(i + 1, 5, aggregate_sentiment_score)
    # appending the derived values to the list for plotting the graphs
    list_positive.append(round(s_positive, 3))
    list_negative.append(round(s_negative, 3))
    list_neutral.append(round(s_neutral, 3))
    list_aggregate_sentiment_score.append(aggregate_sentiment_score)
    # calculating for the final aggregate sentiment score and probability
    total_tweets = total_tweets + len(polarity)
    total_positive = total_positive + s_positive * len(polarity)
    total_negative = total_negative + s_negative * len(polarity)
    total_neutral = total_neutral + s_neutral * len(polarity)
    total_sentiment_score = total_sentiment_score + aggregate_sentiment_score
    # performing heirarchical clustering on the each of the supply chain events
    X = array(get_polarity_subjectivity_list(polarity, subjectivity))
    ms = MeanShift()
    ms.fit(X)
    labels = ms.labels_
    cluster_centers = ms.cluster_centers_
    n_clusters_ = len(np.unique(labels))
    # now saving it in the "hierarchical_clustering_data.xls" file
    pol_x = (cluster_centers[0][0] + cluster_centers[1][
        0]) / 2  # applying the coordinate geometry centre of two coordinates for the first two cluster points
    sub_y = (cluster_centers[0][1] + cluster_centers[1][1]) / 2
    ws2.write(i + 1, 0, i + 1)
    ws2.write(i + 1, 1, list_sheetnames[i])
    ws2.write(i + 1, 2, pol_x)
    ws2.write(i + 1, 3, sub_y)
    ws2.write(i + 1, 4, n_clusters_)
    # writing all the cluster points
    result_point = ""
    for k in range(n_clusters_):
        result_point = result_point + " ( " + str(round(cluster_centers[k][0], 3)) + " , " + str(
            round(cluster_centers[k][1], 3)) + " )"

    ws2.write(i + 1, 5, result_point)
    # now plotting the hierarchical clustering with the cluster points
    colors = 10 * ['r.', 'g.', 'b.', 'c.', 'k.', 'y.', 'm.']
    for j in range(len(X)):
        plt.plot(X[j][0], X[j][1], colors[labels[j]], markersize=10)

    plt.scatter(cluster_centers[:, 0], cluster_centers[:, 1], marker='x', color='k', s=150, linewidths=5, zorder=10)
    plt.title(list_sheetnames[i])
    plt.xlabel("Polarity---------------------->")
    plt.ylabel("Subjectivity------------------>")
    plt.show()
    # now building the dendrogram for each of the supply chain incidents
    temp = []
    score = []
    for j in range(len(polarity)):
        temp = []
        temp.append(polarity[j])
        temp.append(subjectivity[j])
        score.append(temp)

    Y = pdist(score)
    Z = linkage(Y, 'ward')
    c, coph_dists = cophenet(Z, Y)  # c contains the coorelative distance for the clusters
    ws2.write(i + 1, 6, round(c, 3))
    print("" + list_sheetnames[i] + ": " + str(c))
    # calculating the full dednrogram
    plt.figure(figsize=(25, 10))
    plt.title("Dendrogram : " + list_sheetnames[i])
    plt.xlabel('sample index')
    plt.ylabel('distance')
    fancy_dendrogram(
        Z,
        truncate_mode='lastp',
        p=12,
        leaf_rotation=90.,
        leaf_font_size=12.,
        show_contracted=True,
        # annotate_above=10,
        max_d=1.5
    )
    plt.show()

wb.save("Aggregate_Sentiment_Score.xls")  # saving the aggregate sentiment score data
wb2.save("hierarchical_clustering_data.xls")  # saving the hierarchical clustering data

# now plotting the bar graphs for positive, negative, neutral probablity for the incidents and
#  aggregate sentiment score with the corresponding the incidents

# now plotting the aggregate positive  and negative sentiment
#  for each of the supply chain incidents on the same bar graph
N = total_num_sheets
ind = np.arange(N)
width = 0.15
fix, ax = plt.subplots()
pos1_bar = ax.bar(ind, list_positive, width, color='y')

# adding the negative values to the same bar graph
neg1_bar = ax.bar(ind + width, list_negative, width, color='r')

ax.set_ylabel('Probability')
ax.set_title('Sentiment Probability for Supply Chain Events')
ax.set_xticks(ind + width / 2)
ax.set_xticklabels(get7charsfromlist(list_sheetnames))
ax.legend((pos1_bar[0], neg1_bar[0]), ('Positive', 'Negative'))


def autolabel(rects):
    """
    Attach a text label above each bar displaying its height
    """
    for rect in rects:
        height = rect.get_height()
        ax.text(rect.get_x() + rect.get_width() / 2., 1.05 * height,
                '%f' % height,
                ha='center', va='bottom')


# autolabel displays the value in the bar graph as shown
autolabel(pos1_bar)
autolabel(neg1_bar)
plt.show()

# now plotting the bar graph for the aggregate sentiment score with the supply chain incidents
plt.barh(ind, list_aggregate_sentiment_score, align='center', alpha=0.5)
plt.yticks(ind, list_sheetnames)
plt.xlabel('Supply Chain Incidents------------------------------>')
plt.ylabel('Aggregate Sentiment Score----------------------->')
plt.title("Aggregate Sentiment Score vs. Supply Chain Incidents")
plt.show()

# now doing the visulization in the form of geo-mapping
# UNCOMMENT the below codes,  it to get the location of the users in the "preprocessed.xls" file and get the "geographic_data_v2.xls" file
# the program needs the "geographic_data_v2.xls file to plot the maps for each of the supply chain incidents.
g = geocoders.GoogleV3()
'''''

extractor = twitter_setup()
for i in range(total_num_sheets):
    sheet_read = pd.read_excel("preprocessed.xls", list_sheetnames[i])
    user_date = sheet_read["User & Date"].values.tolist()
    sentiment_score = sheet_read["Polarity"].values.tolist()
    ws = wb.add_sheet(list_sheetnames[i])
    ws.write(0, 0, "Tweet No. ", style_blue)
    ws.write(0, 1, "Twitter User", style_blue)
    ws.write(0, 2, "Sentiment Value", style_blue)
    ws.write(0, 3, "Place", style_blue)
    ws.write(0, 4, "Latitude", style_blue)
    ws.write(0, 5, "Longitude", style_blue)
    k = 1
    for j in range(len(user_date)):
        s = getuser(user_date[j])
        try:
            tweets = extractor.user_timeline(screen_name=s, count=2)
            place = g.geocode(tweets[0].user.location)
            print(s + " " + tweets[0].user.location + " Lat: " + str(place.latitude) + " Lon: " + str(place.longitude))
            ws.write(k, 0, j + 1)
            ws.write(k, 1, s)
            ws.write(k, 2, sentiment_score[j])
            ws.write(k, 3, tweets[0].user.location)
            ws.write(k, 4, place.latitude)
            ws.write(k, 5, place.longitude)
            k = k + 1

        except:
            pass

wb.save("geographic_data_v2.xls")

'''

# once the file with the locations "geographic_data_v2.xls" file is created, we can start plotting the points on the
# map red marker denotes that it has the negative sentiment value, where as blue and green denotes the neutral and
# positive sentiment value respectively also, if it has cloud sign it means it has strong sentiment, and if it has
# the info sign it means it has less strong sentiment whether positive or negative

# create a map , creating markers on the map, blue marker denotes the neutral sentiment, red marker-> negative
# sentiment and green marker-> positive sentiment this_map = folium.map(prefer_canvas=True)

list_sheetnames_geo = pd.ExcelFile("geographic_data_v2.xls").sheet_names
total_num_sheets_geo = len(list_sheetnames_geo)

# creating list to contain all the latitudes, longitudes, places and the sentiments
latitudes_list = []
longitudes_list = []
places_list = []
sentiments_list = []

for i in range(total_num_sheets_geo):
    sheet_read = pd.read_excel("geographic_data_v2.xls", list_sheetnames_geo[i])
    sentiments = sheet_read["Sentiment Value"].values.tolist()
    latitudes = sheet_read["Latitude"].values.tolist()
    longitudes = sheet_read["Longitude"].values.tolist()
    places = sheet_read["Place"].values.tolist()
    # appending the values to the defined list
    for k in range(len(sentiments)):
        if True:
            sentiments_list.append(sentiments[k])
            latitudes_list.append(latitudes[k])
            longitudes_list.append(longitudes[k])
            places_list.append(places[k])

    # now making a data-frame with markers to show on the map
    data = pd.DataFrame({
        'lat': latitudes,
        'lon': longitudes,
        'name': places,
        'sentiment': sentiments,
    })
    # now make an empty map
    m = folium.Map(location=[20, 0], zoom_start=2)
    # popup will be used when a particular marker will be clicked,
    # it will display the sentiment value along with the corresponding place

    for j in range(0, len(data)):
        try:
            if data.iloc[j]['sentiment'] > 0:
                folium.Marker([data.iloc[j]['lat'], data.iloc[j]['lon']],
                              popup="Sentiment :  " + str(round(data.iloc[j]['sentiment'], 3)) + " \nLocation :" + str
                              (data.iloc[j]['name']),
                              icon=folium.Icon(color='green')).add_to(m)
            elif data.iloc[j]['sentiment'] < 0:
                folium.Marker([data.iloc[j]['lat'], data.iloc[j]['lon']],
                              popup="Sentiment :  " + str(round(data.iloc[j]['sentiment'], 3)) + " \nLocation: " + str
                              (data.iloc[j]['name']),
                              icon=folium.Icon(color='red')).add_to(m)
            else:
                folium.Marker([data.iloc[j]['lat'], data.iloc[j]['lon']],
                              popup="Sentiment :  " + str(round(data.iloc[j]['sentiment'], 3)) + " \nLocation : " + str
                              (data.iloc[j]['name']),
                              icon=folium.Icon(color='blue')).add_to(m)
        except:
            # print("error"+str(j))
            pass

    m.save(list_sheetnames_geo[i] + "_geo.html")

# now plotting all the points in the single "geo_result_supply_chain_events.html" file

data = pd.DataFrame({
    'lat': latitudes_list,
    'lon': longitudes_list,
    'name': places_list,
    'sentiment': sentiments_list
})

m_all = folium.Map(location=[20, 0], zoom_start=2)  # zoom value =2 will show the globe fitting the screen initially
for j in range(0, len(data)):
    try:
        if data.iloc[j]['sentiment'] > 0:
            folium.Marker([data.iloc[j]['lat'], data.iloc[j]['lon']],
                          popup="Sentiment :  " + str(round(data.iloc[j]['sentiment'], 3)) + " \nLocation :  " + str
                          (data.iloc[j]['name']),
                          icon=folium.Icon(color='green')).add_to(m_all)
        elif data.iloc[j]['sentiment'] < 0:
            folium.Marker([data.iloc[j]['lat'], data.iloc[j]['lon']],
                          popup="Sentiment :  " + str(round(data.iloc[j]['sentiment'], 3)) + " \nLocation:  " + str
                          (data.iloc[j]['name']),
                          icon=folium.Icon(color='red')).add_to(m_all)
        else:
            folium.Marker([data.iloc[j]['lat'], data.iloc[j]['lon']],
                          popup="Sentiment :  " + str(round(data.iloc[j]['sentiment'], 3)) + " \nLocation:  " + str
                          (data.iloc[j]['name']),
                          icon=folium.Icon(color='blue')).add_to(m_all)

    except:
        pass

m_all.save("geovisualization_supply_chain_events.html")

exit(0)
