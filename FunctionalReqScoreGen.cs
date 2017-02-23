using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApplication2
{
    class FunctionalReqScoreGen
    {
        private Excel.Workbook currentWorkBook = null;
        private Excel.Application currentExcelApp = null;
        private Excel.Worksheet currentSheet = null;
        private Excel.Worksheet scorersSheet = null;
        private Excel.Range scorersRange = null;
        private Excel.Range initialRange = null;
        private System.Array Scorers;
        private System.Array MyValues;
        private String filePath = "";

        private int endRow;
        private int startRow = 3;
        private int startCol = 1;
        private int numberingColumn = 1;
        private int titleColumn = 2;

        private int noSystems = 0;
        int noScorers = 0;
        private String upperRange = "";
        private String lowerRange = "";
        private String sheetName = "";


        List<String> scorerNames = new List<String>();

        public int tasks = 16;
        public int tasksCompleted = 0;

        private int exitCode = -1;
        public bool appFinishedRunning = false;

        TableFactory tf;
        private bool tfRunning = true;



        public FunctionalReqScoreGen(String filePath, int noScorers, int startRow, int startCol)
        {
            this.filePath = filePath;
            this.noScorers = noScorers;

            if (startCol > 0 && startRow > 0)
            {
                this.startRow = startRow;
                this.startCol = startCol;
            }
        }

        public double getCompletionProgress()
        {
            double taskCompleted;
            if (tf == null)
                return 0;
            else
            {
                if (!tfRunning)
                    taskCompleted = this.tasksCompleted;
                else
                    tasksCompleted = tf.getTasksCompleted();
                return tasksCompleted / (double)tasks;
            }
        }

        public int getExitCode()
        {
            return exitCode;
        }

        public Boolean validateExcelFile()
        {
            try
            {
                Console.WriteLine("Setting up excel file...");
                /*
                currentExcelApp = new Excel.Application();
                currentExcelApp.Visible = true;
                currentWorkBook = currentExcelApp.Workbooks.Open(filePath); // "C:\\Users\\Alan\\Documents\\GitHub\\Lumenia\\PracticeWorkbooks\\Book4.xlsx"
                currentSheet = (Excel.Worksheet)currentWorkBook.Sheets[1];*/

                
                
                initialRange = currentSheet.get_Range(this.lowerRange, this.upperRange);
                MyValues = initialRange.Cells.Value;
                Console.WriteLine("Excel file successfully set up");
                return true;
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                Console.WriteLine(e.StackTrace);
                return false;
            }
        }

        public Boolean validateRangeSelected()
        {
            //String g = initialRange.Validation.ToString();
            //Console.WriteLine(g);
            return true;
        }

        public int findNoScorers()
        {
            scorersSheet = (Excel.Worksheet)currentWorkBook.Sheets["Scorers"];
            scorersRange = scorersSheet.get_Range("A1", "A500");
            Scorers = scorersRange.Cells.Value;

            int currentRow = startRow;
            int currentCol = startCol;
            int noScorers = 0;
            Boolean endOfDocumentFound = false;

            while (!endOfDocumentFound)
            {
                if (Scorers.GetValue(currentRow, currentCol) == null)
                    endOfDocumentFound = true;
                else
                {
                    String val = Scorers.GetValue(currentRow, currentCol).ToString();
                    if (val != "")
                    {
                        scorerNames.Add(val);
                        //Console.WriteLine("Found " + val);
                        noScorers++;
                    }

                }
                currentRow++;
            }
            return noScorers;
        }


        public void run()
        {
            // Initial testing

            tf = new TableFactory(filePath, startRow, startCol, tasksCompleted);
            tfRunning = true;
            tf.makeTable();
/*
            Console.WriteLine("Application Finished Running");
            exitCode = 0;
            appFinishedRunning = true;
        }
            
       public void test()
       { */
            tfRunning = false;

            tasksCompleted = tf.getTasksCompleted();

            sheetName = tf.getNewSheetName();
            lowerRange = tf.getStartRange();
            upperRange = tf.getEndRange();
            currentSheet = tf.getNewSheet();
            noSystems = tf.getNumberOfSystems();

            Console.WriteLine("Sheet Name = " + tf.getNewSheetName() + " Table range = " + tf.getStartRange() + ":" + tf.getEndRange());
            //return 8;
            tasksCompleted++;
            //this.noScorers = findNoScorers();
            

            if (noSystems < 1)
            {
                Console.WriteLine("##############################\n");
                Console.WriteLine("Error: No/Invalid Number of Systems Assigned");
                Console.WriteLine("Current Number of Systems Assigned: " + this.noSystems);
                Console.WriteLine("\n##############################\n");
                appFinishedRunning = true;
                exitCode = -1;
            }

            else if (noScorers < 1)
            {
                Console.WriteLine("##############################\n");
                Console.WriteLine("Error: No/Invalid Number of Scorers Assigned");
                Console.WriteLine("Current Number of Scorers Assigned: " + this.noScorers);
                Console.WriteLine("\n##############################\n");
                appFinishedRunning = true;
                exitCode = -1;
            }

            else if (!validateExcelFile())
            {
                Console.WriteLine("##############################\n");
                Console.WriteLine("Error: Problem Opening file at the path assigned");
                Console.WriteLine("Current File Path: " + this.filePath);
                Console.WriteLine("\n##############################\n");
                appFinishedRunning = true;
                exitCode = -1;
            }

            else if (!validateRangeSelected())
            {
                Console.WriteLine("##############################\n");
                Console.WriteLine("Error: No/Invalid range assigned");
                Console.WriteLine("Current Range: (" + lowerRange + ":" + upperRange + ")");
                Console.WriteLine("\n##############################\n");
                appFinishedRunning = true;
                exitCode = -1;
            }
            else
            {

                Console.WriteLine("All tests passed...\n");
                tasksCompleted++;
                //return 0;

                int currRow = 2;//startRow + 1;

                numberingColumn = 1;// startCol;
                titleColumn = 2;// numberingColumn + 1;


                Boolean endOfDocumentFound = false;

                try
                {
                    while (!endOfDocumentFound)
                    {
                        if ((MyValues.GetValue(currRow, numberingColumn) == null) && (MyValues.GetValue(currRow, titleColumn) == null))
                            endOfDocumentFound = true;
                        else
                            currRow++;
                    }

                }
                catch (System.IndexOutOfRangeException ex)
                {
                    string err = ex.StackTrace;
                }

                endRow = currRow - 1;
                currRow = 2;// startRow + 1;
                Console.WriteLine("starting row = " + startRow);
                Console.WriteLine("End row = " + endRow);
                Console.WriteLine("curr row = " + currRow);
                tasksCompleted++;
                //return 4;
                List<Heading> headings = new List<Heading>();
                Heading currentHeading = null;
                SubHeading currentSubHeading = null;
                SubSubHeading currentSubSubHeading = null;
                Requirement currentRequirement = null;


                bool headingInUse = false;
                bool subHeadingInUse = false;
                bool subSubHeadingInUse = false;

                Console.WriteLine("Finding all headings in sheet...");
                // get headings in the sheet
                while (currRow != endRow)
                {
                    Location currentLocation = new Location(currRow + startRow - 1, numberingColumn + startCol - 1);
                    String title = "";
                    if (MyValues.GetValue(currRow, titleColumn) != null)
                        title = MyValues.GetValue(currRow, titleColumn).ToString();

                    if (MyValues.GetValue(currRow, numberingColumn) != null)
                    {

                        string cellData = MyValues.GetValue(currRow, numberingColumn).ToString();
                        int count = decimalPlacesCounter(cellData);
                        if (count == 0)
                        {
                            //Console.WriteLine("heading found '" + cellData + "'");
                            currentHeading = new Heading(cellData, currentLocation, title);  // create new heading object from the current location
                            headings.Add(currentHeading);

                            headingInUse = true;
                            subHeadingInUse = false;
                            subSubHeadingInUse = false;
                        }
                        else if (count == 1)
                        {
                            //Console.WriteLine("Subheading found '" + cellData + "'");
                            currentSubHeading = new SubHeading(cellData, currentLocation, title);  // create new heading object from the current location
                            currentHeading.addSubHeadingToList(currentSubHeading);

                            headingInUse = false;
                            subHeadingInUse = true;
                            subSubHeadingInUse = false;
                        }
                        else if (count == 2)
                        {
                            //Console.WriteLine("SubSubHeading found '" + cellData + "'");
                            currentSubSubHeading = new SubSubHeading(cellData, currentLocation, title);  // create new heading object from the current location
                            currentSubHeading.addSubSubHeadingToList(currentSubSubHeading);

                            headingInUse = false;
                            subHeadingInUse = false;
                            subSubHeadingInUse = true;
                        }
                        else if (count == 3)
                        {
                            //Console.WriteLine("Requirement found '" + cellData + "'");
                            //currentSubSubHeading = new SubSubHeading2(cellData, currentLocation, title);  // create new heading object from the current location
                            //currentSubHeading.addSubSubHeadingToList(currentSubSubHeading);
                        }
                    }
                    else if (MyValues.GetValue(currRow, titleColumn) != null)
                    {
                        string cellData = MyValues.GetValue(currRow, titleColumn).ToString();
                        Console.WriteLine("Cell on row " + currRow + " col " + numberingColumn + " = null" + " : Celldata in title colloim is " + cellData + " col " + titleColumn);
                        if (isRequirement(cellData))
                        {
                            currentRequirement = new Requirement(cellData, currentLocation, title);
                            int count = decimalPlacesCounter(cellData);
                            if (subSubHeadingInUse)
                            {
                                Console.WriteLine("subsub requirement found '" + cellData + "'");
                                currentSubSubHeading.addRequirementToList(currentRequirement);
                            }
                            else if(subHeadingInUse)
                            {
                                Console.WriteLine("sub requirement found '" + cellData + "'");
                                currentSubHeading.addRequirementToList(currentRequirement);
                            }
                            else if(headingInUse)
                            {
                                Console.WriteLine("heading requirement found '" + cellData + "'");
                                currentHeading.addRequirementToList(currentRequirement);
                            }
                        }
                        else if(cellData == "Average") {
                        if(subSubHeadingInUse) {
                                subSubHeadingInUse = false;
                                subHeadingInUse = true;
                        }
                        else if(subHeadingInUse) {
                                subHeadingInUse = false;
                                headingInUse = true;
                            }
                        }


                    }

                    currRow++;
                }
                tasksCompleted++;

                Console.Write("Calculating averages...");
                // Calculate Heading Averages
                for (int systemNumber = 1; systemNumber <= noSystems; systemNumber++)
                {
                    
                    for (int i = 0; i < headings.Count; i++)
                    {
                        currentHeading = headings[i];
                        List<SubHeading> subHeadings = currentHeading.getSubHeadings();
                        String currentHeadingAverage = "";

                        


                        for (int j = 0; j < subHeadings.Count; j++)
                        {
                            currentSubHeading = subHeadings[j];
                            Location currentSubHeadingLocation = new Location(currentSubHeading.getLocation().getRow(), (currentSubHeading.getLocation().getColumn() + 1 + systemNumber));
                            String currentSubHeadingAverage_ssh = ""; // to accumlate average for number of subsubheadings
                            String currentSubHeadingAverage_reqs = ""; // to accumlate average for number of requirements
                            Boolean averageAlreadyAssigned = false;


                            if (currentSubHeading.hasSubSubHeadings() && currentSubHeading.hasRequirements())
                            {
                                List<JointHeading> j_list = tf.buildJointHeadingList(currentSubHeading);
                                List<Location> locations = new List<Location>();
                                int ssh_count = 0;
                                int req_count = 0;

                                for(int k = 0;k <j_list.Count;k++)
                                {
                                    JointHeading jh = j_list[k];
                                    if(jh.isSubSubHeading()) {
                                        currentSubSubHeading = jh.getSubSubHeading();
                                        int ssh_row = currentSubSubHeading.getLocation().getRow() + currentSubSubHeading.getRequirements().Count + 1;
                                        int ssh_col = currentSubSubHeading.getLocation().getColumn() + 1 + systemNumber;
                                        Location ssh_l = new Location(ssh_row, ssh_col);
                                        locations.Add(ssh_l);
                                        String ssh_data = currentSubSubHeading.assignAverageForRequirements(systemNumber);
                                        writeToSingleCell(ssh_l, ssh_data, 0);
                                        ssh_count++;
                                    }
                                    if(jh.isRequirement()) {
                                        int ssh_row = jh.getLocation().getRow();
                                        int ssh_col = jh.getLocation().getColumn() + 1 + systemNumber;
                                        Location ssh_l = new Location(ssh_row, ssh_col);
                                        locations.Add(ssh_l);
                                        req_count++;
                                    }
                                    
                                }

                                String averageOfRequirements = "";
                                for (int k = 0; k < locations.Count; k++)
                                {
                                    Location newLoc = locations[k];
                                    if (k == (locations.Count - 1))
                                    {

                                        averageOfRequirements = "AVERAGE(" + averageOfRequirements + newLoc.getExcelAddress() + ") ";
                                    }
                                    else
                                    {
                                        averageOfRequirements = averageOfRequirements + newLoc.getExcelAddress() + ", ";
                                    }

                                }
                                averageOfRequirements = "=IFERROR(" + averageOfRequirements + "/10,\"\")";

                                int row = currentSubHeading.getLocation().getRow();// + j_list.Count + 1 + ssh_count;
                                int col = currentSubHeading.getLocation().getColumn() + 1 + systemNumber;
                                Location l = new Location(row, col);
                                String data = averageOfRequirements;
                                writeToSingleCell(l, data, 2);
                                averageAlreadyAssigned = true;
                            }

                            else if (currentSubHeading.hasSubSubHeadings())
                            {
                                List<SubSubHeading> subSubHeadings = currentSubHeading.getSubSubHeadings();
                                for (int k = 0; k < subSubHeadings.Count; k++)
                                {
                                    currentSubSubHeading = subSubHeadings[k];
                                    int row = currentSubSubHeading.getLocation().getRow() + currentSubSubHeading.getRequirements().Count + 1;
                                    int col = currentSubSubHeading.getLocation().getColumn() + 1 + systemNumber;
                                    Location l = new Location(row, col);

                                    String data = currentSubSubHeading.assignAverageForRequirements(systemNumber);
                                    writeToSingleCell(l, data, 0);

                                    if (isValidRow(row) && currentSubSubHeading.hasRequirements())
                                    {
                                        if (k == (subSubHeadings.Count - 1))
                                            currentSubHeadingAverage_ssh = currentSubHeadingAverage_ssh + l.getExcelAddress();
                                        else
                                            currentSubHeadingAverage_ssh = currentSubHeadingAverage_ssh + l.getExcelAddress() + " , ";
                                    }
                                }
                                if (currentSubHeadingAverage_ssh != "")
                                {
                                    currentSubHeadingAverage_ssh = "AVERAGE(" + currentSubHeadingAverage_ssh + ")";
                                    currentSubHeadingAverage_ssh = "=IFERROR(" + currentSubHeadingAverage_ssh + "/10,\"\")";

                                }
                            }

                            else if (currentSubHeading.hasRequirements())
                            {
                                
                                int row = currentSubHeading.getLocation().getRow() + currentSubHeading.getRequirements().Count + 1;
                                int col = currentSubHeading.getLocation().getColumn() + 1 + systemNumber;
                                Location l = new Location(row, col);
                                String data = currentSubHeading.assignAverageForRequirements(systemNumber);
                                writeToSingleCell(l, data, 0);

                                if(currentSubHeading.hasSubSubHeadings())
                                    currentSubHeadingAverage_reqs = data; //"=IFERROR(" + l.getExcelAddress() + "/10,0)";
                                else
                                    currentSubHeadingAverage_reqs = "=IFERROR(" + l.getExcelAddress() + "/10,0)";
                            }


                            String currentSubHeadingAverage = "";

                            if (currentSubHeading.hasSubSubHeadings())
                            {
                                currentSubHeadingAverage = currentSubHeadingAverage_ssh;
                            }

                            if (currentSubHeading.hasRequirements())
                            {
                                currentSubHeadingAverage = currentSubHeadingAverage_reqs;
                            }

                            if (currentSubHeading.hasSubSubHeadings() && currentSubHeading.hasRequirements())
                            {
                                //Console.WriteLine("************************************");
                                //Console.WriteLine("Special Case. It has both requirements and subsubheadings");
                                //Console.WriteLine("String is now: " + currentSubHeadingAverage_1 + " && " + currentSubHeadingAverage_2);
                                //Console.WriteLine("************************************");

                                currentSubHeadingAverage = formatNewAverageString(currentSubHeadingAverage_ssh, currentSubHeadingAverage_reqs);
                                //Console.WriteLine("String is now: " + currentSubHeadingAverage);
                                //Console.WriteLine("************************************");
                            }


                            //else
                            //    Console.WriteLine("Sub heading " + currentSubHeading.getValue() + " has no  reqs");

                            if (!averageAlreadyAssigned)
                            {
                                Location subHeadingAverageLocation = new Location(currentSubHeading.getLocation().getRow(), (currentSubHeading.getLocation().getColumn() + 1 + systemNumber));
                                writeToSingleCell(subHeadingAverageLocation, currentSubHeadingAverage, 0);
                            }


                            //     && (currentSubHeading.hasRequirements() || currentSubHeading.hasSubSubHeadings()))

                            if (j == (subHeadings.Count - 1))
                            {
                                if(currentSubHeading.hasRequirements() || currentSubHeading.hasSubSubHeadings())
                                    currentHeadingAverage = currentHeadingAverage + currentSubHeadingLocation.getExcelAddress();
                                else  // if we are at the last sub heading in a heading, and that sub heading contains no reqs or subsubheadings, remomve the comma ',' from the end of the string
                                {
                                    string[] arr = currentHeadingAverage.Split();
                                    int length = arr.Length;
                                    if(arr[length - 1] == ",")
                                        arr[length - 1] = "";
                                    currentHeadingAverage = "";
                                    for (int idx = 0;idx < arr.Length; idx++){
                                        String s = arr[idx];
                                        currentHeadingAverage += s;
                                    }
                                   
                                }
                            }
                            else
                                if (currentSubHeading.hasRequirements() || currentSubHeading.hasSubSubHeadings())
                                 currentHeadingAverage = currentHeadingAverage + currentSubHeadingLocation.getExcelAddress() + " , ";
                        }

                        if (currentHeading.hasRequirements())
                        {

                            int row = currentHeading.getLocation().getRow() + currentHeading.getRequirements().Count + 1;
                            int col = currentHeading.getLocation().getColumn() + 1 + systemNumber;
                            Location l = new Location(row, col);
                            String data = currentHeading.assignAverageForRequirements(systemNumber);
                            Console.WriteLine("Writing " + data + " to location " + l.getAddress());
                            writeToSingleCell(l, data, 0);
                            currentHeadingAverage += l.getExcelAddress() + "/10"; // divide average of requirements by 10 so we don't get X00%
                        }

                        if (currentHeadingAverage != "")
                        {
                            
                            currentHeadingAverage = "AVERAGE(" + currentHeadingAverage + ")";
                        }
                        currentHeadingAverage = "=IFERROR(" + currentHeadingAverage + ", 0)";
                        Location currentHeadingAverageLocation = new Location(currentHeading.getLocation().getRow(), (currentHeading.getLocation().getColumn() + 1 + systemNumber));
                        //Console.WriteLine("Heading location " + currentHeadingAverageLocation.getExcelAddress());
                        writeToSingleCell(currentHeadingAverageLocation, currentHeadingAverage, 0);
                    }
                }

                tasksCompleted++;

                Console.WriteLine("DONE");

                Console.WriteLine("Setting up scorer sheets...");
                setUpScorerSheets(headings, 2);
                Console.WriteLine("Scorer sheets are set up");
                tasksCompleted++;

                Console.Write("Applying Conditional Formatting...");
                tf.conditionalFormat();
                Console.WriteLine("Done");
                tasksCompleted++;

                Console.WriteLine("Application Finished Running");
                exitCode = 0;
                appFinishedRunning = true;
                
            }
        }

        private string formatNewAverageString(string currentSubHeadingAverage_ssh, string currentSubHeadingAverage_req)
        {
            /*
             * Assumes currentSubHeadingAverage_2 is of the format: =IFERROR(AVERAGE(H171 , H175 , H182 , H188 , H194)/10,"")
             * Assumes currentSubHeadingAverage_1 is of the format: =IFERROR(H163/10,0)
             * 
             */
            String new_1 = "";
            String new_2 = "";

            Console.WriteLine("currSubHead1: " + currentSubHeadingAverage_ssh + "    CurrSubHead2:" + currentSubHeadingAverage_req);

        //   ) ,"")

            //new_2 = currentSubHeadingAverage_2.Replace("=IFERROR(", "");
            new_2 = currentSubHeadingAverage_req.Replace("=IFERROR(AVERAGE(", "");
            new_2 = new_2.Replace(") ,\"\")","");

            //Console.WriteLine("This is new_2: + " + new_2);

            /*
             * currentSubHeadingAverage_1 should now be of the format: AVERAGE(H171 , H175 , H182 , H188 , H194)
             * 
             */

            Console.WriteLine("currSubHead1: " + new_1 + "    CurrSubHead2:" + new_2);


            new_1 = currentSubHeadingAverage_ssh.Replace("=IFERROR(AVERAGE(", "");
            new_1 = new_1.Replace(")/10,\"\")", "");


            Console.WriteLine("currSubHead1: " + new_1 + "    CurrSubHead2:" + new_2);
            //Console.WriteLine("This is new_1: + " + new_1);
            /*
            * currentSubHeadingAverage_1 should now be of the format: H163
             * 
            */


            /*
             * We should then return a string of the format: =IFERROR((AVERAGE(" + H163 + "," + AVERAGE(H171 , H175 , H182 , H188 , H194) + ")/10),0)
             * 
             */
            //return "=IFERROR((AVERAGE(" + new_1 + "," + new_2 + ")/10),0)";

            String ret = "=IFERROR((AVERAGE(" + new_1 + "," + new_2 + ")/10),0)";

            return ret;


        }








        private bool isValidRow(int currentRow)
        {
            int row = currentRow;
            row = row - startRow;
            for (int i = 1; i <= (2 + noSystems); i++)
            {
                if (MyValues.GetValue(row, i) != null)
                {
                    if (MyValues.GetValue(row, i).ToString() == "No score required")
                    {
                        //Console.WriteLine("     No Score required for row " + currentRow);
                        return false;
                    }
                }
            }
            //Console.WriteLine("     row " + currentRow + " is valid");
            return true;
        }

        private void writeToSingleCell(Location cellLocation, string value, int mode)
        {
            int row = cellLocation.getRow();
            int col = cellLocation.getColumn();
            if (mode == 1 || mode == 2)
                Console.WriteLine("Writing to " + col + "  (" + cellLocation.getExcelAddress() + "): " + value);
            if (mode == 0 || mode == 2)
                currentSheet.Cells[row, col] = value;
        }

        private void writeToSingleCell(Location cellLocation, string value, int mode, int cellType, double cellWidth, double cellHeight, bool boldFlag, bool italicsFlag, bool wrapTextFlag, System.Drawing.Color cellColour,
            System.Drawing.Color textColour, Excel.XlHAlign horizAlignment, Excel.XlVAlign vertAlignment, String fontName, int fontSize)
        {
            int row = cellLocation.getRow();
            int col = cellLocation.getColumn();

            if (mode == 1 || mode == 2)
                Console.WriteLine("Writing to " + col + "  (" + cellLocation.getExcelAddress() + "): " + value);
            if (mode == 0 || mode == 2)
            {
                currentSheet.Cells[row, col] = value;
                Excel.Range r = currentSheet.Cells[row, col] as Excel.Range;

                if (cellWidth > 0)
                    currentSheet.Columns[startCol].ColumnWidth = cellWidth;
                if (cellHeight > 0)
                    currentSheet.Rows[startRow].RowHeight = cellHeight;

                r.Interior.Color = System.Drawing.ColorTranslator.ToOle(cellColour);
                r.Font.Color = System.Drawing.ColorTranslator.ToOle(textColour);
                r.Font.Bold = boldFlag;
                r.Font.Italic = italicsFlag;
                r.WrapText = wrapTextFlag;
                r.HorizontalAlignment = horizAlignment;
                r.VerticalAlignment = vertAlignment;
                r.Font.Name = fontName;
                r.Font.Size = fontSize;
            }
        }

        private void writeToMultipleCells(Location startRangeLocation, Location endRangeLocation, string value, int mode, int cellType, double cellWidth, double cellHeight, bool boldFlag, bool italicsFlag, bool wrapTextFlag, System.Drawing.Color cellColour,
           System.Drawing.Color textColour, Excel.XlHAlign horizAlignment, Excel.XlVAlign vertAlignment, String fontName, int fontSize)
        {
            int startRow = startRangeLocation.getRow();
            int startCol = startRangeLocation.getColumn();
            int endRow = endRangeLocation.getRow();
            int endCol = endRangeLocation.getColumn();

            string startRange = getExcelColumnName(startCol) + startRow.ToString();
            string endRange = getExcelColumnName(endCol) + endRow.ToString();

            if (mode == 1 || mode == 2)
                Console.WriteLine("Writing to " + startCol + "  (" + startRange + ":" + endRange + ") " + value);
            if (mode == 0 || mode == 2)
            {

                if(cellWidth > 0)
                    currentSheet.Columns[startCol].ColumnWidth = cellWidth;
                if (cellHeight > 0)
                    currentSheet.Rows[startRow].RowHeight = cellHeight;
                Excel.Range r = currentSheet.get_Range(startRange, endRange); //Excel.Range r = newTableSheet.Cells[startRow, startCol] as Excel.Range;
                r.Interior.Color = System.Drawing.ColorTranslator.ToOle(cellColour);
                r.Font.Color = System.Drawing.ColorTranslator.ToOle(textColour);
                r.Font.Bold = boldFlag;
                r.Font.Italic = italicsFlag;
                r.WrapText = wrapTextFlag;
                r.HorizontalAlignment = horizAlignment;
                r.VerticalAlignment = vertAlignment;
                r.Font.Name = fontName;
                r.Font.Size = fontSize;

                if (value != "" && value != null)
                    r.Cells.Value = value;
            }
        }

        private int decimalPlacesCounter(String value)
        {
            int count = 0;
            char[] arr;

            if (value == null || value == "")
                count = 0;
            else
            {
                arr = value.ToCharArray();
                for (int i = 0; i < arr.Length; i++)
                    if (arr[i] == '.')
                        count++;
            }

            return count;
        }

        private string getColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            var value = "";

            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];

            value += letters[index % letters.Length];

            return value;
        }


        private string getExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public void setUpScorerSheets(List<Heading> headings, int space)
        {

            // first paste the data
            Console.WriteLine("Pasting Data...");
            int columnLength = noSystems + 3 + space;
            int col = startCol;
            int row = 1;// startRow;

            String lowerRange = getExcelColumnName(currentSheet.Range[this.lowerRange].Column);
            String upperRange = getExcelColumnName(currentSheet.Range[this.upperRange].Column);


            for (int scorer = 1; scorer <= noScorers; scorer++)
            {
                col = col + columnLength;
                Console.WriteLine("Pasting to Column: " + col + "\n     Home Range: (" + lowerRange + ":" + upperRange + ")");
                
                
                Excel.Range r1 = currentSheet.get_Range(lowerRange + ":" + upperRange);
                Excel.Range r2 = currentSheet.get_Range(getExcelColumnName(col) + ":" + getExcelColumnName(col));
                try
                {
                    r1.Copy(r2);
                }
                catch (System.Runtime.InteropServices.COMException e)
                {
                    Console.WriteLine(e.StackTrace);
                }

                Location startRange = new Location(row, col);
                Location endRange = new Location(row + 3, col + noSystems + 2);
                writeToMultipleCells(startRange, endRange, "", 0, 0, 0, 0, true, false, true, System.Drawing.Color.LightGray, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignCenter, "Calibri", 16);

                endRange = new Location(row, col);
                writeToSingleCell(endRange, scorer.ToString(), 0, 0, 0, 0, true, false, true, System.Drawing.Color.DarkSlateGray, System.Drawing.Color.White, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignCenter, "Calibri", 16); // write scorer number to make it easier to read on spreadsheet
                endRange = new Location(row, col + 1);
                writeToMultipleCells(startRange, endRange, "", 0, 0, 0, 0, true, false, true, System.Drawing.Color.DarkSlateGray, System.Drawing.Color.White, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignCenter, "Calibri", 16);
                writeToSingleCell(endRange, "FUNCTIONALITY REQUIREMENTS SCORESHEET", 0, 0, 0, 0, true, false, true, System.Drawing.Color.DarkSlateGray, System.Drawing.Color.White, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignCenter, "Calibri", 16);

                startRange = new Location(row, col + 2);
                writeToSingleCell(startRange, "<Enter name>", 0, 0, 0, 0, false, false, false, System.Drawing.Color.LightGray, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignLeft, Excel.XlVAlign.xlVAlignCenter, "Calibri", 16);
                
                //See worksheet "INSTRUCTIONS (REQTS & ITT)" for full instructions on how to use this scoresheet
                endRange = new Location(row + 2, col + 1);
                writeToSingleCell(endRange, "See worksheet \"INSTRUCTIONS (REQTS & ITT)\" for full instructions on how to use this scoresheet", 0, 0, 0, 0, true, false, true, System.Drawing.Color.LightGray, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignLeft, Excel.XlVAlign.xlVAlignBottom, "Calibri", 12);

            }
            Console.WriteLine("Data Pasted");



            Console.WriteLine("Assigning scorer table values to main...");
            for (int i = 0; i < headings.Count; i++)
            {
                Heading heading = headings[i];


                // handle case where a heading has a direct set of requirements
                if(heading.hasRequirements())
                {
                    List<Requirement> reqs = heading.getRequirements();
                    for (int k = 0; k < reqs.Count; k++)
                    {
                        Requirement inReq = reqs[k];
                        int homeRow = inReq.getLocation().getRow();
                        int homeCol = inReq.getLocation().getColumn();
                        homeCol = homeCol + 2;
                        for (int l = 1; l <= noSystems; l++)
                        {

                            String average = "";
                            int indexCol = homeCol + columnLength;

                            for (int m = 1; m <= noScorers; m++)
                            {

                                Location loc = new Location(homeRow, indexCol);
                                if (m != noScorers)
                                    average = average + loc.getExcelAddress() + ", ";
                                else
                                    average = average + loc.getExcelAddress() + ")";
                                indexCol += columnLength;

                            }

                            average = "=IFERROR(AVERAGE(" + average + ", \"\")";
                            Location cell = new Location(homeRow, homeCol);
                            writeToSingleCell(cell, average, 0);
                            homeCol = homeCol + 1;

                            // make string
                        }

                    }
                }


                List<SubHeading> subHeadings = heading.getSubHeadings();
                for (int j = 0; j < subHeadings.Count; j++)
                {
                    SubHeading subHeading = subHeadings[j];
                    if (subHeading.hasRequirements())
                    {
                        List<Requirement> reqs = subHeading.getRequirements();
                        for (int k = 0; k < reqs.Count; k++)
                        {
                            Requirement inReq = reqs[k];
                            int homeRow = inReq.getLocation().getRow();
                            int homeCol = inReq.getLocation().getColumn();
                            homeCol = homeCol + 2;
                            for (int l = 1; l <= noSystems; l++)
                            {

                                String average = "";
                                int indexCol = homeCol + columnLength;

                                for (int m = 1; m <= noScorers; m++)
                                {

                                    Location loc = new Location(homeRow, indexCol);
                                    if (m != noScorers)
                                        average = average + loc.getExcelAddress() + ", ";
                                    else
                                        average = average + loc.getExcelAddress() + ")";
                                    indexCol += columnLength;

                                }

                                average = "=IFERROR(AVERAGE(" + average + ", \"\")";
                                Location cell = new Location(homeRow, homeCol);
                                writeToSingleCell(cell, average, 0);
                                homeCol = homeCol + 1;

                                // make string
                            }

                        }
                    }
                    if (subHeading.hasSubSubHeadings())
                    {
                        List<SubSubHeading> subSubHeadings = subHeading.getSubSubHeadings();
                        for (int w = 0; w < subSubHeadings.Count; w++)
                        {
                            SubSubHeading req = subSubHeadings[w];
                            if (req.hasRequirements())
                            {
                                List<Requirement> reqs = req.getRequirements();
                                for (int k = 0; k < reqs.Count; k++)
                                {
                                    Requirement inReq = reqs[k];
                                    int homeRow = inReq.getLocation().getRow();
                                    int homeCol = inReq.getLocation().getColumn();
                                    homeCol = homeCol + 2;
                                    for (int l = 1; l <= noSystems; l++)
                                    {

                                        String average = "";
                                        int indexCol = homeCol + columnLength;

                                        for (int m = 1; m <= noScorers; m++)
                                        {

                                            Location loc = new Location(homeRow, indexCol);
                                            if (m != noScorers)
                                                average = average + loc.getExcelAddress() + ", ";
                                            else
                                                average = average + loc.getExcelAddress() + ")";
                                            indexCol += columnLength;

                                        }

                                        average = "=IFERROR(AVERAGE(" + average + ", \"\")";
                                        Location cell = new Location(homeRow, homeCol);
                                        writeToSingleCell(cell, average, 0);
                                        homeCol = homeCol + 1;

                                        // make string
                                    }

                                }
                            }
                        }
                    }
                }
            }

            Console.WriteLine("Scorer table values assigned");
        }


        private Boolean isRequirement(String value)
        {
            int n;
            bool isNumeric = false;
            isNumeric = int.TryParse(value, out n);
            if (isNumeric)
                return true;
            else
            {
                char[] arr = value.ToCharArray();

                for (int i = 0; i < arr.Length; i++)
                {

                    switch (arr[i])
                    {
                        case '1':
                            return true;
                        case '2':
                            return true;
                        case '3':
                            return true;
                        case '4':
                            return true;
                        case '5':
                            return true;
                        case '6':
                            return true;
                        case '7':
                            return true;
                        case '8':
                            return true;
                        case '9':
                            return true;
                    }
                }
                return isNumeric;

            }
        }

        // #############################################################
        // ##############################################################

        


    }

}
