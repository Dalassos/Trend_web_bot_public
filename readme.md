A web scrapper to visit Honeywell Trend BEMS controller's webpages and modify information on a large scale. 

Trend Web Bot:

This tool can be used to visit and scrape information from multiple Honeywell Trend Controllers webpages, as well as change information in them.
This is intended to facilitate site maintenance and can be quite useful for housekeeping before a front-end upgrade.

Example of use:
-scan a whole site for vCNC conflict when you lose comms	
-scan a whole site to find out if you have any BBMD devices setup	
-scan all of your sites to know where your time masters are 	
-modify time server for time sync in all your time masters at once	
-scan all your sites for failure in alarm destinations	
-modify alarm destination IP address in all your controllers at once	
-scan all your devices for GUID at once to identify conflict and sort it in an excel sheet for easy comparison	
-turn on switch W1 in all your controllers at once	

***********************************
limitations:
Has been tested with IQ3, IQ4 and IQ5 only. 
Logins will all be tried one after another until access is granted for EVERY controller where a password/PIN is required. If all your controllers have different login, this may be very long.

***********************************
Instructions:

Setup:

The tool relies on the controllers list for which an empty template had been provided; this needs to be filled in with at least the IPs to do anything at all. 
If any or your controllers are PIN/password protected (they should) then all the possible login/password/PIN combinations need to be entered in the login sheet for which an empty template is provided.

If you have a few controllers, then filling the sheet manually once may be the easiest option. Likely, you have a lot of controller and it makes more sense to create a sheet from your front-end. You should be able to export this from IQ Vision/Niagara front-end through clever BQL scripting, although I did not test this method yet. From 963, the easiest way is to export the list from your database. You can do this with "sql server management studio" *SSMS* (free ms application) by copying your 963 database (i96X_data fron your server's drive, you will need to stop 963 to copy and paste it), attaching it to a local SQL server with SSMS and running the following SQL query to get the equivalent OS_full_list excel sheet:

SELECT s.[siteID],s.[siteLabel],s.[siteConnectionString],
	l.[LanNo],l.[theLabel],
	 os.[NodeAddress],os.[theLabel],os.[outPin],os.[deviCeResponse],os.[nodeIpAddr],
	 os.[TimeMasterStatus]
  FROM [i96X].[dbo].[Outstations] AS os
FULL OUTER JOIN [i96X].[dbo].[Lans] AS l ON os.[lanID]=l.[lanID]
FULL OUTER JOIN [i96X].[dbo].[SiteDetails] AS s ON s.[siteID]=l.[siteID]
ORDER BY s.[siteLabel]


Use: 

Extract somewhere where you can access all your controllers IPs, double click on trend_web_bot.py. You will need to select sites (from the cotroller list at setup) and properties to be extracted each time you run the software. By default, Scan/Replace will scan only. If you open the Replace interface and check "Replace", then you will also be replacing info. The interface should be self explanatory if you get there. Your results will be in an excel sheet "scan_result_datetime.xls" in the same folder. IF you get any error, make sure you filled in "OS_full_list" and "OS_logins", kept their original name and kept them in their original location (main folder).

***THIS WILL NOT WORK WITHOUT SETING UP THE OS_full_list EXCEL SHEET***
