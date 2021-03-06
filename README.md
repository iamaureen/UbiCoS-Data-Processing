# UbiCoS-Data-Processing

* **wordcount_avg.py** </br>
input: excel file with all user utterances (sheet 0), utterances from Khan Academy (sheet 1), and modelbook i.e., image and general message uttrances (sheet 2) </br>
output: average word per utterance in Khan Academy, average word per utterance in Modelbook

The text files used in the following codes are in specific format obtained from the UbiCoS server. 
File Names may vary with different studies, will have to change the names accordingly in the python script

* **gm_data.py** </br>
input: gm_data.txt </br>
[run ubicos server, and hit http://127.0.0.1:8000/getGeneralChatMsg in the browser.
It will display a dict with users as key, and their comments in general chat as values for that key] </br>
output: an excel file 'gm.xls' with users as one column, and their comments on the second column. 

* **image_data.py** </br>
input: image_data.txt </br>
[run ubicos server, and hit http://127.0.0.1:8000/getimageCommentMsg in the browser.
It will display a dict with users as key, and their comments in galleries as values for that key]</br>
output: an excel file 'image.xls' with users as one column, and their comments on the second column.

* **merge.py:** </br>
input: gm_data.txt, image_data.txt, and ka_data.txt </br>
[run ubicos server, and hit http://127.0.0.1:8000/getkhanAcademyMsg in the browser.
It will display a dict with users as key, and their comments in galleries as values for that key] </br>
output: an excel file 'merge.xls' with three different sheets: general chat, image comments and Khan Academy comments.
For all sheets, users as one column, and their comments on the second column. For Khan Academy sheet, we also have whether
each utterance/comment is of type answer/question 

* **samesheet.py** </br>
input: merge.xls created from merge.py </br>
output: an excel file 'all.xls' all the values combined in one sheet instead of three different ones

* **rowdata.py** </br>
input: merge.xls created from merge.py </br>
output: an excel file 'row.xls' all the values combined in one sheet, and put them into rows sequentially instead of columns
(this is the final version, we use this one to get the final excel file)

**So, steps:**
-----------
1. run server and get data from the server: </br>
general chat message: http://127.0.0.1:8000/getGeneralChatMsg </br>
image messages (all galleries): http://127.0.0.1:8000/getimageCommentMsg </br>
khan academy messages: http://127.0.0.1:8000/getkhanAcademyMsg

2. for each of the urls above, save the data in three separate text files.

3. modify names in merge.py as needed and run merge.py

4. run rowdata.py


