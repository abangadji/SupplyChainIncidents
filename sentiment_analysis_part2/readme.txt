Sentiment Analysis of the Supply Chain Incidents- Part II
The task is divided into 3 parts:

1. Calculation of the aggregate sentiment score

2. Hierarchical Clustering /Making dendrograms of the Supply Chain Incidents

3. Geo-Graph visualization of each of the supply chain incidents


1. AGGREGATE SENTIMENT SCORE: In this two kinds of calculation are done.
                              1. Probability of positive, negative, neutral sentiments are calculated for each of the supply chain incidents.
							  2. Aggregate Sentiment Value from polarity and subjectivity value as : summation(polarity * objectivity)
							      or, summation(polarity *(1 - subjectivity))
								 The value got from summation is a value used to compare the sentiment score of the supply chain incidents.
		                         Bar-graph for probabiliy, and aggregate sentiment is shown in the screenshots folder.



2.	HIERARCHICAL CLUSTERING & DENDROGRAMS: The hierarchical clustering is done for each of the supply chain incidents (x,y), where x = polarity, y = subjectivity
                              
							    The number of clusters is determined programmatically. The equivalent (x,y) ie. polarity, subjectivity is determined using the mean -point formula of coordinate geometry.
								The value for each of the events is saved in another excel file "hierarchical_clustering_data.xls". 
								
								Dendrograms is made with the pre-calculated clusters and coorelative-distance is calculated. This distance is saved in the same excel file.
								Graphs of clustering and dendrograms are shown in screenshots folder.
								


3.  GEO-GRAPHING VISUALIZATION: First of all, twitter handle of all users is calculated. 
                                Using geocoder api, the latitude and longitude, and location of all tweets are found and saved in another excel file, "geographic_data_v2.xls".
								Using folium package of python, the latitude and longitude are marked on the map for each of the supply chain incidents.
								The red marker denotes a negative sentiment on the map. Similiarly, blue denotes neutral and green denotes positive.
								Finally, a map is made combining all the sentiment from all the events. and plotted.
								The analysis is in the output/analysis_report_2.doc file.
								The map is stored in the html file. Use browser to open it.
								
		