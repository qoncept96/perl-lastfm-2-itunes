Script is designed to update information on number of replays, and also on the date of a last replay in the library of **iTunes** based on the data received from **Last.fm**.

Algoritm:
  * Connect to the Last.fm service and download of statistic on music replay by the user for the whole time of using the service.
  * Analysis of songs in the library of iTunes and update of inforamation on the number of replays and date of a last replay.

Requirements:
  * **Windows XP**+
  * **ActivePerl**
  * **XML::DOM** module for **ActivePerl**
  * **iTunes**