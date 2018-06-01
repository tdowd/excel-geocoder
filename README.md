# excel-geocoder
A VBA application for geocoding and reverse geocoding in Excel. Supports both Google's free and enterprise for business geocoder (Maps API for Work).


# Background
GIS and geospatial data science applications will usually require geocoding of locations or reverse geocoding of latitude/longitude at some point in the analysis. While most of this analysis is frequently done in something a little more involved than Excel (python, arcGIS, ESRI, etc.), sometimes I have found that doing something quick with a dataset in Excel can be more efficient than working with a SQL DB or creating a python script.

This project was influenced by a blog post by josephglover on his blog [Police Analyst](http://policeanalyst.com/using-the-google-geocoding-api-in-excel/). josephglover's module on accessing the free Google geocoder was the foundation which I used to make the reverse geocoder and to add flexibility to use Google's Maps API for Work Enterprise geocoder.


# Prerequisites
* Enable developer tab in Excel. Instructions from MSFT can be [found here](https://msdn.microsoft.com/en-us/library/bb608625.aspx).
* Within the VB IDE, go to setup VBA module and run the AddReference() Sub


# Installation
* Import the .bas file into your project.
* To use Google's Maps API for Work geocoder, view the code in the VB IDE and change the `gintType` constant equal to `1` and insert your Google Client ID and Google Secret Key into the `gstrClientID` and `gstrKey` constants respectively.
* To use Google's API Premium Plan, change the `gintType` constant equal to `2` and insert your API key into the `gstrKey` constant.
* To use Google's Free Geocoding API, change the `gintType` constant equal to `0`.


# Usage
* `=AddressGeocode(address)`
	* Takes in the address of the location we want to geocode and returns the first latitude, longitude pair from the geocoder.
* `=ReverseGeocode(lat,long)`
	* Takes in a latitude, longitude pair and returns the first formatted address from the geocoder.


# TODO
* Test cases
* Functionality for Bing Maps, Data Science Toolkit, etc.
* Fix for forcing too many requests at one time
