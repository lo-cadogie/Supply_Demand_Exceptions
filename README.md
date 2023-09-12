# Supply_Demand_Exceptions

Python Code
  Data files are downloaded from Oracle (MRP) and saved in an 'Input' Folder
  Data files are loaded into a Pandas dataframe
  Inventory values in not-nettable inventory locations are isolated using ".isin(list)" and removed using "drop"
  Buyer names are renamed using .map
  Other data cleaning methods used are: del, .drop, .rename, display, .unique, .dtypes, .fillna, .sort_values, groupby, .apply(lambda x: x.cumsum()) 
  Once all updates are made, data files are merged together, written, saved and opened in an single Excel file within an 'Output' Folder.
  
Excel File
  Data from the output file is copied and pasted into the zzExceptions.xlsm Template
  Hitting the 'Fix Colors' button runs the VBA script
  VBA color codes: Reschedule outs, cancels and Reschedule Ins
  Conditional formatting color codes negative inventory and run out dates in red
  If the Num Column is filtered on a number, the "Back" and "Next" buttons can be used to navigate through each item

** Item names, using assemblies, planner codes, buyer names and suppliers were changed to provide this example **
  
