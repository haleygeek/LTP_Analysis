# LTP_Analysis

Python 2.7 program to analyze electrophysiological data from a long-term potentiation experiment. This program assumes raw data is in an xlsx file as follows. Row 1= Experiment Name, Rows 2-31 = Raw slopes of baseline, Rows 32+ = Raw slopes of post-induction. Assumes no empty cells in dataset, but may change that.

Data need to be in an excel .xlsx workbook 

Sheet containing raw slopes from LTP (or LTD) experiment should be formatted as:

  1 Column = 1 Experiment
  
    For each column:
    
      Row 1 = Experiment name
      
      Row 2-31 = Baseline raw slopes 
      
      Row 32+ = Post-induction raw slopes
      

There should be no empty cells (will change this later)

Number of rows is not limited

Number of columns is not limited

Other than row 1, if a cell is found to be empty (before max_row is reached) or to contain text, the program terminates and tells you to check your spreadsheet.

User must know the path and filename, as well as the name of the sheet within the workbook. Will change this later.

Output is printed to the screen for convenience, but saved in a new sheet within the existing workbook as "Norm Sheet Name"
If the sheet already exists, the new sheet will be "Norm Sheet Name 1"

