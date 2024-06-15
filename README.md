Application/User interface for tracking and storing data for multiple end users daily work.


Description of used functions:

**initUI**
  declraed layout , logic for underlying User Interface

**initDB**
  initialize database connection into UI, loads db structure

**loadtabledata**
  loads data for QTableWidget

**your_data**
  set a filter to QtableWidget db data based on person selected in Case entered by window 

**StartDate**
  sets start date after pressing star date button

**holdtime**
  sets hold time after pressing holdtime button

**FinishDate**
  sets finish time after pressing FinishDate button

**createNewCase**
  add new record to database and Qtablewidget  

**save**
   save selected record to database and Qtablewidget    

**delete**
   delete selected record to database and Qtablewidget

**double click**
  connect to database and retreive information from it , and next show all data into Qtablewidget and in labels 

**get_text**
  supporting function for double click that handles propper formating of fields and layout 

**SelectNewCase**
  similar to double click but triggers only when create new case button is triggered 

**filter**
  Silter records in Qtablewidget according open cases and your name

**Filter off**
  turns off all filters

**calc_Total_time**
  calculate time that past since presing start date till finish date 

  
  
