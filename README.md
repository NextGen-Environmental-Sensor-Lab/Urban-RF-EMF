<img src="/images/rfemf_nyc_googlemaps_all.png" alt="RF-EMF" width="8000">

# Urban-RF-EMF
Urban RF-EMF measurements in NYC. 
Data and code repo

### Measuring the Spatial Distribution of Environmental Radio Frequency Electromagnetic Field Exposure in New York City

Arno Thielens, Ricardo Toledo-Crow, Salvatore Davi, Sassama Hema  
Advanced Science Research Center, City University of New York.

<img src="/images/expom-rf-4.jpg" alt="ExpoM RF-4" width="300">
<img src="/images/IMG_3772 2.jpeg" alt="Sal and Sassama on their way..." width="500">


The data was acquired with an ExpoM RF field sensor in two periods. It is split into the following sets:
+ #1 fall 2024  : outdoor paths including measurements on the water ferry
+ #3 spring 2025 : repeat of the outdoor paths of #1
+ indoor measurements : a few indoor paths done during #1
+ train measurements : a few train (NYC MTA) measurements during #1

Data acquisition paths were done in the five boroughs of NYC, in different environments
+ **Boroughs**: M, BK, Q, BX, SI, FERRY (Manhattan, Brooklyn, Queens, Bronx, Staten Island, Ferry)
+ **Environments** : C, R, G, I, T (Commercial, Residential, Greenery, Indoors, Transport)

Folders: 
+ **ExpoM_data**: Source data files from the ExpoM RF sensor in tab separated value format.
+ **Excel_data**: Processed CSV files to Excel files with csv2excel_batch.py, cleaned up some. File names were manually edited. See below note for naming.
+ **Excel_data_aggregated**: Excel files with bands aggregated by technology (broadcast, cellular upload, cellular download, WLAN, TDD).
+ **Excel_inventory**: Summary of totals by technology per path for season 1, 3, indoors and trains. Made with script make_inventory_totals_bis_bis.py. See script header for more.
+ **GoogleEarth_kml**: 'My Places.kmz' for Google Earth Pro with all measurement paths in season 1. Also some cool images from GoogleEarth.
+ **Heatmaps**: Heatmap type display of all paths by bands and aggregations for season 1. Also Python scripts to make them.
+ **QGIS_layers**: QGIS file with all the information: census block population (cbp), census block expanded population (ecbp), pedestrian mobility ranking (pm), all rf-emf paths in season 1.
+ **Season1-Season3_correlation**: Excel file with Spearman Rank Correlation analysis for measurememnts in the two seasons (and Wilcoxon bias).
+ **Population_Pedestrian_analysis**: Excel file with Spearman Rank Correlation analysis of measurememnts in season 1: RF-EMF exposure to population and to foot traffic in the city (pedestrian mobility ranking).

NOTES: The names of the excel files were manually edited to include the categories info (environment, borough) The general format is: 
#### *YYYY-MM-DD_hh.mm.ss E B location.xlsx* where *E* = environment, *B* = borough. 
Season 3 is often refered to as Season 2 in the publication (Season3=Season2).

