About the two python files:

automated\_analysis.py (currently prism\_xml\_method.py (in case I forget to change it later )) and final\_result.py



The script determines if you're using a mac or windows and proceeds accordingly



**Important!!!** 

Change the path to files and folders. Set Batch True or False

Change batch path 

You might need bead result folder (I may have created this already)

**Important!!!**





**Automated analysis.py** expects a folder with pulling data. It processes the first ten pulling files greater than 89KB (changeable).

* It copies this into excel in specific bead regions (data input sheet).
* For each excel bead sheet it prints certain region(that should always have data) and the value of K1 (which should be  OK to show the data is not short) for error checking. 
* Then, it copies the values of T and U columns for each bead and puts it in prism (ensure your prism path and version is put accurately)
* Then it runs Automated datcol(a prism script) which exports bead.txt for all beads... These are the results of the prism analysis. Each bead.txt has values of k0, k1, tau, viscosity*. This refreshes on each call to prevent stale data.* 

  * *Note:* Prism prefers complete data and will stop if 10 beads data is not there. However, the bead data until the missing bead will have results. Thus, I restricted the analysis to not proceed if there's less than 8 pulling data. 
* After the k0, k1, tau, viscosity values are put in excel, **if the values are really large, the tab for that bead will be coloured orange. if the values/bead data is missing the tab is coloured purple** and print the bead numbers with valid data**.**
* The next is the trimming of Data Output to MATLAB1 to the row of the shortest column and the appending 'Data Output to MATLAB1' and 'Data Output to MATLAB2' to file 2 and 1 respectively. 
* MATLAB runs (ensure path is correct) and produces Allbeads.csv (A caveat to this is not able to see the graph produced by MATLAB for debug reasons. I may end up saving maybe one image to make thing better.)
* The result of allbeads is put into the input from MATLAB sheet. The result is available in Graphpad Graph\&Stats sheets.
* To make the final result step easier, the result of G'0.1Hz, G'1Hz, G'10Hz are copied into Gp\_values csv





**Final\_results.py** aims to append each folders results to a single prism file. Change the directory for each folder. As far as the summary data for excel is there, it should work



