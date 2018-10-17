# excel-geocoder
A VBA application for geocoding and reverse geocoding in Excel. Supports both Google's free and enterprise for business geocoder (Maps API for Work).


# Background
GIS and geospatial data science applications will usually require geocoding of locations or reverse geocoding of latitude/longitude at some point in the analysis. While most of this analysis is frequently done in something a little more involved than Excel (python, arcGIS, ESRI, etc.), sometimes I have found that doing something quick with a dataset in Excel can be more efficient than working with a SQL DB or creating a python script.

This project was influenced by a blog post by josephglover. josephglover's module on accessing the free Google geocoder was the foundation which I used to make the reverse geocoder and to add flexibility to use Google's Maps API for Work Enterprise geocoder.


# Prerequisites
* Enable developer tab in Excel. Instructions from MSFT can be [found here](https://msdn.microsoft.com/en-us/library/bb608625.aspx).
* Within the VB IDE, add "Microsoft XML, v6.0" as a Reference. Can be found within *Tools* - *References*.


# Installation
* Import the .bas file into your project.
* To use Google's Maps API for Work geocoder, view the code in the VB IDE and change the `gintType` constant equal to `1` and insert your Google Client ID and Google Secret Key into the `gstrClientID` and `gstrKey` constants respectively.
* To use Google's API Premium Plan, change the `gintType` constant equal to `2` and insert your API key into the `gstrKey` constant.
* To use Google's Free Geocoding API, change the `gintType` constant equal to `0`.
* Note that as of late Summer/early Fall 2018, Google's Free Geocoding API now also requires a key. More information can be [found here](https://developers.google.com/maps/documentation/geocoding/usage-and-billing) and [here](https://developers.google.com/maps/documentation/geocoding/get-api-key). Still set your `gintType` to `0` and insert your new API key into the `gstrKey` constant and you should still be able to geocode in Excel.


# Usage
* `=AddressGeocode(address)`
	* Takes in the address of the location we want to geocode and returns the first latitude, longitude pair from the geocoder.
* `=ReverseGeocode(lat,long)`
	* Takes in a latitude, longitude pair and returns the first formatted address from the geocoder.


# TODO
* Clean up code around key management as a result of Google's 2018 key management changes
* Test cases
* Functionality for Bing Maps, Data Science Toolkit, etc.
* Fix for forcing too many requests at one time
