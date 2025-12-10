<img src="/GoogleEarth_kml/manhattan 3d.jpg" alt="RF-EMF" width="300">

# Urban-RF-EMF
Urban RF-EMF measurements in NYC. 
Data and code repo

### Measuring the Spatial Distribution of Environmental Radio- Frequency Electromagnetic Field Exposure in New York City

Arno Thielens, Ricardo Toledo-Crow, Salvatore Davi, Sassama Hema  
Advanced Science Research Center, City University of New York.

<img src="/images/expom-rf-4.jpg" alt="ExpoM RF-4" width="300">
<img src="/images/IMG_3772 2.jpeg" alt="Sal and Sassama on their way..." width="300">


The data was acquired with an ExpoM RF field sensor in two periods. It is split into the following sets:
+ #1 fall 2024  : outdoor paths including measurements on the water ferry
+ #3 spring 2025 : repeat of the outdoor paths of #1
+ indoor measurements : a few indoor paths done during #1
+ train measurements : a few train (NYC MTA) measurements during #1

Data acquisition paths were done in the five boroughs of NYC, in different environments
+ boroughs : M, BK, Q, BX, SI, FERRY (Manhattan, Brooklyn, Queens, Bronx, Staten Island, Ferry)
+ environments : C, R, G, I, T (Commercial, Residential, Greenery, Indoors, Transport)

Folders: 
+ ExpoM_data : has source data files from the ExpoM RF sensor in tab separated value format
+ Excel_data : processed files with csv2excel_batch.py, cleaned up some. Names were manually edited. See below note.
+ GoogleEarth_kml: has 'My Places.kmz' for Google Earth Pro with all measurement paths in season 1, census blocks and pedestrian mobility lines. Also some cool images form GEP

Scripts: running on Anaconda Python 3.13.5 in Visual Sudio Code
+ csv2excel_batch.py 
+ make_inventory_totals_bis.py : generates the inventory excel files 
+ excel2image.py : makes heatmap images of the individual path files (?)

Inventory Excel Files : these files are created with the make_inventory_totals_bis.py script. See script header for more

NOTE: the names of the excel files were manually edited to include the categories info (environment, borough)

The general format is "YYYY-MM-DD_hh.mm.ss E B location.xlsx" E=environment, B=borough
