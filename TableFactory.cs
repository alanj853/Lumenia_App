﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApplication2
{
    class TableFactory
    {
        private const int TYPE_IGNORE = -1;
        private const int TYPE_DEFAULT = 0;
        private const int TYPE_HEADING = 1;
        private const int TYPE_SUBHEADING = 2;
        private const int TYPE_SUBSUBHEADING2 = 3;
        private const int TYPE_SUBSUBHEADING = 4;
        private const int TYPE_REQUIREMENT = 5;
        private const int TYPE_AVERAGE = 6;
        private const int TYPE_NOSCORE = 7;
        private const int TYPE_SYSTEM = 8;
        private const int TYPE_COMMENT = 9;

        private Excel.Workbook currentWorkBook = null;
        private Excel.Application currentExcelApp = null;
        private Excel.Worksheet requirementsSheet = null;
        private Excel.Worksheet systemsSheet = null;
        private Excel.Worksheet newTableSheet = null;
        private System.Array MyValues;
        private System.Array MySystems;
        private System.Array newTableValues;
        private Excel.Range initialRange = null;
        private Excel.Range systemsRange = null;
        private Excel.Range newTableRange = null;
        private String filePath = "";

        private int endRow = 1;
        private int startRow = 1; // start row for data in Systems and Requirements Sheets
        private int startCol = 1;

        private int newStartRow = 1; // start row for data in new sheet being built, default = 1
        private int newStartCol = 1;

        int noSystems = 0;
        private int lengthOfNewSheet = 0;
        Boolean sheetSuccessfullySetUp = false;
        List<String> systemNames = new List<String>();
        List<Location> virtualSpreadsheet = new List<Location>();
        List<int> alreadyWrittenTo = new List<int>();

        int tasksCompleted = 0;


        public TableFactory(String filePath, int newStartRow, int newStartCol, int tasksCompleted)
        {
            this.filePath = filePath;
            this.newStartRow = newStartRow;
            this.newStartCol = newStartCol;
            this.tasksCompleted = tasksCompleted;

            alreadyWrittenTo.Add(-1);
        }


        public int getNumberOfSystems()
        {
            return noSystems;
        }

        public String getStartRange()
        {
            int col = newStartCol;
            int row = newStartRow;
            if (sheetSuccessfullySetUp)
                return getExcelColumnName(col) + row.ToString();

            return "Sheet has not been set up yet.";
        }

        public String getEndRange()
        {
            int col = newStartCol + noSystems + 2;
            int row = lengthOfNewSheet;

            if (sheetSuccessfullySetUp)
                return getExcelColumnName(col) + row.ToString();

            return "Sheet has not been set up yet.";
        }

        public Excel.Worksheet getNewSheet()
        {

            return newTableSheet;

        }

        public int getTasksCompleted()
        {
            return tasksCompleted;
        }

        public String getNewSheetName()
        {
            if (sheetSuccessfullySetUp)
                return newTableSheet.Name;

            return "Sheet has not been set up yet.";
        }

        public Boolean validateExcelFile()
        {
            try
            {
                currentExcelApp = new Excel.Application();
                currentExcelApp.Visible = true;
                currentWorkBook = currentExcelApp.Workbooks.Open(filePath); // "C:\\Users\\Alan\\Documents\\GitHub\\Lumenia\\PracticeWorkbooks\\Book4.xlsx"
                currentWorkBook.Windows[1].WindowState = Excel.XlWindowState.xlMinimized;
                Console.Write("Setting up excel file...");

                requirementsSheet = (Excel.Worksheet)currentWorkBook.Sheets["Requirements"];
                systemsSheet = (Excel.Worksheet)currentWorkBook.Sheets["Systems"];

                initialRange = requirementsSheet.get_Range("A1", "C10000");
                systemsRange = systemsSheet.get_Range("A1", "A500");


                MyValues = initialRange.Cells.Value;
                MySystems = systemsRange.Cells.Value;

                Console.WriteLine("DONE");
                tasksCompleted++;
                return true;
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                Console.WriteLine(e.StackTrace);
                return false;
            }
        }

        public int makeTable()
        {
            if (!validateExcelFile())
            {
                Console.WriteLine("Could not validate Excel file");
                return -1;
            }

            int currRow = startRow;
            int numberingColumn = startCol;
            int titleColumn = numberingColumn + 1;


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

            endRow = currRow;
            currRow = startRow;//startRow + 1;
            Console.WriteLine("starting row = " + startRow);
            Console.WriteLine("End row = " + endRow);
            Console.WriteLine("curr row = " + currRow);


            List<Heading> headings = new List<Heading>();
            Heading currentHeading = null;
            SubHeading currentSubHeading = null;
            SubSubHeading currentSubSubHeading = null;
            Requirement currentRequirement = null;


            bool subHeadingInUse = false;

            Console.Write("Finding all headings in sheet...");
            // get headings in the sheet
            while (currRow != endRow)
            {
                Location currentLocation = new Location(currRow, numberingColumn);
                String title = "";
                if (MyValues.GetValue(currRow, titleColumn) != null)
                    title = MyValues.GetValue(currRow, titleColumn).ToString();

                if (MyValues.GetValue(currRow, numberingColumn) != null)
                {

                    string cellData = MyValues.GetValue(currRow, numberingColumn).ToString();
                    int y = decimalPlacesCounter(cellData);
                    if (y == 0)
                    {
                        //Console.WriteLine("heading found '" + cellData + "'");
                        currentHeading = new Heading(cellData, currentLocation, title);  // create new heading object from the current location
                        headings.Add(currentHeading);
                        subHeadingInUse = false;
                    }
                    else if (y == 1)
                    {
                        //Console.WriteLine("Subheading found '" + cellData + "'");
                        currentSubHeading = new SubHeading(cellData, currentLocation, title);  // create new heading object from the current location
                        currentHeading.addSubHeadingToList(currentSubHeading);
                        subHeadingInUse = true;
                    }
                    else if (y == 2)
                    {
                        if (MyValues.GetValue(currRow, titleColumn) != null)
                        {
                            //Console.WriteLine("SubSubHeading found '" + cellData + "'");
                            currentSubSubHeading = new SubSubHeading(cellData, currentLocation, title);  // create new heading object from the current location
                            currentSubHeading.addSubSubHeadingToList(currentSubSubHeading);
                            subHeadingInUse = false;
                        }
                        else
                        {
                           // Console.WriteLine("Requirement found '" + cellData + "'");
                            currentRequirement = new Requirement(cellData, currentLocation, "");
                            currentRequirement.setUpdateTitle();
                            currentSubHeading.addRequirementToList(currentRequirement);
                        }
                    }
                    else if (y == 3)
                    {
                        //Console.WriteLine("Requirement found '" + cellData + "'");
                        //currentSubSubHeading2 = new SubSubHeading2(cellData, currentLocation, title);  // create new heading object from the current location
                        //currentSubHeading.addSubSubHeadingToList(currentSubSubHeading2);
                        //Console.WriteLine("Requirement found '" + cellData + "'");
                        currentRequirement = new Requirement(cellData, currentLocation, "");
                        currentRequirement.setUpdateTitle();
                        if (subHeadingInUse)
                            currentSubHeading.addRequirementToList(currentRequirement);
                        else
                            currentSubSubHeading.addRequirementToList(currentRequirement);
                    }
                }
                else if (MyValues.GetValue(currRow, titleColumn) != null)
                {
                    string cellData = MyValues.GetValue(currRow, titleColumn).ToString();
                    if (isRequirement(cellData))
                    {
                        currentLocation = new Location(currRow, titleColumn);
                        currentRequirement = new Requirement(cellData, currentLocation, "");
                        // Console.WriteLine("Requirement found '" + cellData + "'");
                        if (subHeadingInUse)
                            currentSubSubHeading.addRequirementToList(currentRequirement);
                        else
                            currentSubHeading.addRequirementToList(currentRequirement);
                    }


                }
                //Console.WriteLine("Row = " + currRow);
                currRow++;
            }
            Console.WriteLine("DONE");
            tasksCompleted++;

            Console.Write("Finding number of systems...");
            noSystems = findNoSystems();
            Console.WriteLine("DONE\n   Found " + noSystems + " systems");
            tasksCompleted++;

            for (int i = 0; i < headings.Count; i++)
            {
                //Console.WriteLine("Heading " + (i + 1) + ": " + headings[i].getValue());
                currentHeading = headings[i];
                List<SubHeading> subHeadings = currentHeading.getSubHeadings();
                for (int j = 0; j < subHeadings.Count; j++)
                {
                    currentSubHeading = subHeadings[j];
                    //Console.WriteLine("     SubHeading " + (j + 1) + ": " + currentSubHeading.getValue());
                    List<SubSubHeading> subSubHeadings = currentSubHeading.getSubSubHeadings();
                    for (int k = 0; k < subSubHeadings.Count; k++)
                    {
                        currentSubSubHeading = subSubHeadings[k];
                        //Console.WriteLine("         Requirement " + (k + 1) + ": " + currentSubSubHeading.getValue());
                    }
                }
            }

            Console.Write("Building 'Functional Requirements Scores' Sheet...");
            try
            {

                newTableSheet = (Excel.Worksheet)currentWorkBook.Sheets.Add(After: currentWorkBook.Sheets["Systems"]);
                newTableSheet.Name = "Functional Requirements Scores";//"Table";
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                //Console.WriteLine(ex.ErrorCode + ": " + ex.ToString());
            }

            newTableRange = newTableSheet.get_Range("A1", "Z10000");
            newTableValues = newTableRange.Cells.Value;

            //int rowIndex = 1;
            //int colIndex = 1;

            int rowIndex = newStartRow;
            int colIndex = newStartCol;

            for (int i = 0; i < headings.Count; i++)
            {
                // filling in the first row with headings
                if (i == 0)
                {
                    Location loc = new Location(1, colIndex);
                    //writeToSingleCell(loc, systemNames[i], 0, 8, 15, 55, true, false, true, System.Drawing.Color.DeepSkyBlue, System.Drawing.Color.White, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignTop, "Arial Narrow", 9);

                    colIndex = newStartCol + 2;

                    for (int j = 0; j < noSystems; j++)
                    {
                        loc = new Location(rowIndex, colIndex);
                        writeToSingleCell(loc, systemNames[j], 0, 8, 15, 55, true, false, true, System.Drawing.Color.DeepSkyBlue, System.Drawing.Color.White, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignTop, "Arial Narrow", 9);
                        colIndex++;
                    }

                    loc = new Location(rowIndex, colIndex);
                    writeToSingleCell(loc, "Comments", 0, 9, 35, 55, true, false, true, System.Drawing.Color.DeepSkyBlue, System.Drawing.Color.White, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignCenter, "Arial Narrow", 11);
                    colIndex = newStartCol;
                    rowIndex++;
                }
                Heading h = headings[i];
                Location headingNumberLocation = new Location(rowIndex, colIndex);
                Location headingTitleLocation = new Location(rowIndex, colIndex + 1);
                String headingNumber = h.getValue();
                String headingTitle = h.getTitle();

                writeToSingleCell(headingNumberLocation, headingNumber, 0, 1, 5, 21, true, false, true, System.Drawing.Color.Black, System.Drawing.Color.White, Excel.XlHAlign.xlHAlignLeft, Excel.XlVAlign.xlVAlignCenter, "Arial", 10);
                writeToSingleCell(headingTitleLocation, headingTitle, 0, 1, 63, 21, true, false, true, System.Drawing.Color.Black, System.Drawing.Color.White, Excel.XlHAlign.xlHAlignLeft, Excel.XlVAlign.xlVAlignCenter, "Arial", 10);

                rowIndex++;

                List<SubHeading> subHeadings = h.getSubHeadings();
                for (int j = 0; j < subHeadings.Count; j++)
                {
                    SubHeading s = subHeadings[j];
                    Location subHeadingNumberLocation = new Location(rowIndex, colIndex);
                    Location subHeadingTitleLocation = new Location(rowIndex, colIndex + 1);
                    String subHeadingNumber = s.getValue();
                    String subHeadingTitle = s.getTitle();

                    //writeToSingleCell(subHeadingNumberLocation, subHeadingNumber, 0);
                    //writeToSingleCell(subHeadingTitleLocation, subHeadingTitle, 0);

                    writeToSingleCell(subHeadingNumberLocation, subHeadingNumber, 0, TYPE_SUBHEADING, 5, 21, true, false, true, System.Drawing.Color.DarkGray, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignLeft, Excel.XlVAlign.xlVAlignTop, "Calibri", 11);
                    writeToSingleCell(subHeadingTitleLocation, subHeadingTitle, 0, TYPE_SUBHEADING, 63, 21, true, false, true, System.Drawing.Color.DarkGray, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignGeneral, Excel.XlVAlign.xlVAlignTop, "Calibri", 11);

                    rowIndex++;

                    Boolean noScoreRequiredFlag_forSubHeading = false; // flag to let program know if it has to assign a "No Score" row... uses different procedure
                   // Boolean noScoreRequiredFlag_forSubSubHeading = false; // flag to let program know if it has to assign a "No Score" row... uses different procedure

                    if (subHeadingNumber.Contains("_x") || subHeadingNumber.Contains("_X"))
                    {
                        noScoreRequiredFlag_forSubHeading = true;
                        if (subHeadingNumber.Contains("_x"))
                            subHeadingNumber = subHeadingNumber.Replace("_x", "");
                        else
                            subHeadingNumber = subHeadingNumber.Replace("_X", "");

                        Location noScoreLocation = new Location(rowIndex, colIndex + 1);
                        writeToSingleCell(subHeadingNumberLocation, subHeadingNumber, 0, TYPE_SUBHEADING, 5, 21, true, false, true, System.Drawing.Color.DarkGray, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignLeft, Excel.XlVAlign.xlVAlignTop, "Calibri", 11);
                        writeToSingleCell(noScoreLocation, "No Score Required", 0, TYPE_NOSCORE, 63, 21, true, false, true, System.Drawing.Color.AliceBlue, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignTop, "Calibri", 11);
                        rowIndex++;
                    }

                    if (s.hasRequirements() && !noScoreRequiredFlag_forSubHeading)
                    {
                        List<Requirement> reqs = s.getRequirements();
                        for (int k = 0; k < reqs.Count; k++)
                        {
                            Requirement req = reqs[k];
                            Location reqNumberLocation = new Location(rowIndex, colIndex + 1);
                            String reqNumber = req.getTitle();

                            writeToSingleCell(reqNumberLocation, reqNumber, 0, TYPE_REQUIREMENT, 63, 21, true, false, true, System.Drawing.Color.AliceBlue, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignTop, "Calibri", 11);
                            rowIndex++;
                        }
                        Location reqAverageRow = new Location(rowIndex, colIndex + 1);
                        writeToSingleCell(reqAverageRow, "Average", 0, TYPE_AVERAGE, 63, 21, false, true, true, System.Drawing.Color.LightBlue, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignBottom, "Arial", 10);
                        rowIndex++;
                    }
                    if (s.hasSubSubHeadings())
                    {
                        List<SubSubHeading> subSubHeadings = s.getSubSubHeadings();
                        for (int k = 0; k < subSubHeadings.Count; k++)
                        {
                            SubSubHeading subSubHeading = subSubHeadings[k];
                            Location subSubHeadingNumberLocation = new Location(rowIndex, colIndex);
                            Location subSubHeadingTitleLocation = new Location(rowIndex, colIndex + 1);
                            String subSubHeadingNumber = subSubHeading.getValue();
                            String subSubHeadingTitle = subSubHeading.getTitle();

                            //writeToSingleCell(reqNumberLocation, reqNumber, 0);
                            //writeToSingleCell(reqTitleLocation, reqTitle, 0);

                            writeToSingleCell(subSubHeadingNumberLocation, subSubHeadingNumber, 0, TYPE_SUBSUBHEADING, 5, 21, true, false, true, System.Drawing.Color.LightGray, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignLeft, Excel.XlVAlign.xlVAlignTop, "Calibri", 11);
                            writeToSingleCell(subSubHeadingTitleLocation, subSubHeadingTitle, 0, TYPE_SUBSUBHEADING, 63, 21, true, false, true, System.Drawing.Color.LightGray, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignGeneral, Excel.XlVAlign.xlVAlignTop, "Calibri", 11);
                            rowIndex++;

                            /*
                             * THIS IS THE START OF THE NE STUFF ADDED 21/07/16 22:28
                             * 
                             */

                            Boolean noScoreRequiredFlag_forSubSubHeading = false; // flag to let program know if it has to assign a "No Score" row... uses different procedure

                            if (subSubHeadingNumber.Contains("_x") || subSubHeadingNumber.Contains("_X"))
                            {

                                noScoreRequiredFlag_forSubSubHeading = true;
                                if (subSubHeadingNumber.Contains("_x"))
                                    subSubHeadingNumber = subSubHeadingNumber.Replace("_x", "");
                                else
                                    subSubHeadingNumber = subSubHeadingNumber.Replace("_X", "");


                                Location noScoreLocation = new Location(rowIndex, colIndex + 1);
                                writeToSingleCell(subSubHeadingNumberLocation, subSubHeadingNumber, 0, TYPE_SUBSUBHEADING, 5, 21, true, false, true, System.Drawing.Color.DarkGray, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignLeft, Excel.XlVAlign.xlVAlignTop, "Calibri", 11);
                                writeToSingleCell(noScoreLocation, "No Score Required", 0, TYPE_NOSCORE, 63, 21, true, false, true, System.Drawing.Color.AliceBlue, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignTop, "Calibri", 11);
                                rowIndex++;
                            }

                            if (subSubHeading.hasRequirements() && !noScoreRequiredFlag_forSubSubHeading)
                            {
                                List<Requirement> reqs = subSubHeading.getRequirements();
                                for (int l = 0; l < reqs.Count; l++)
                                {
                                    Requirement req = reqs[l];
                                    Location reqNumberLocation = new Location(rowIndex, colIndex + 1);
                                    String reqNumber = req.getTitle();


                                    writeToSingleCell(reqNumberLocation, reqNumber, 0, TYPE_REQUIREMENT, 63, 21, true, false, true, System.Drawing.Color.AliceBlue, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignTop, "Calibri", 11);
                                    rowIndex++;
                                }

                            }

                            // original before change on 21/07/16
                            /*if (subSubHeading.hasRequirements())
                            {
                                List<Requirement> reqs = subSubHeading.getRequirements();
                                for (int l = 0; l < reqs.Count; l++)
                                {
                                    Requirement req = reqs[l];
                                    Location reqNumberLocation = new Location(rowIndex, colIndex + 1);
                                    String reqNumber = req.getTitle();


                                    writeToSingleCell(reqNumberLocation, reqNumber, 0, TYPE_REQUIREMENT, 63, 21, true, false, true, System.Drawing.Color.AliceBlue, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignTop, "Calibri", 11);
                                    rowIndex++;
                                }

                            }*/

                            if (!noScoreRequiredFlag_forSubSubHeading)
                            {
                                Location reqAverageRow = new Location(rowIndex, colIndex + 1);
                                writeToSingleCell(reqAverageRow, "Average", 0, TYPE_AVERAGE, 63, 21, false, true, true, System.Drawing.Color.LightBlue, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignBottom, "Arial", 10);
                                rowIndex++;
                            } 
                            
                        }
                    }


                }
            }
            Console.WriteLine("DONE");
            tasksCompleted++;
            lengthOfNewSheet = rowIndex - 1;

            //Console.WriteLine("New length = " + lengthOfNewSheet);
            endOfDocumentFound = false;
            colIndex = newStartCol + 2;

            //List<int> alreadyWrittenTo = new List<int>();
            //alreadyWrittenTo.Add(-1);

            Console.Write("Applying Borders...");

            for (int i = 0; i < virtualSpreadsheet.Count; i++)
            {
                Location l = virtualSpreadsheet[i];
                //if (!alreadyWrittenTo.Contains(l.getRow()))
                //   {
                //alreadyWrittenTo.Add(l.getRow());

                if (l.getCellType() == TYPE_HEADING)
                {
                    //Console.WriteLine("ROW: " + l.getRow() + " Heading");
                    //Console.WriteLine("     Adding Data to " + colIndex);
                    Location startRange = new Location(l.getRow(), colIndex);
                    Location endRange = new Location(l.getRow(), colIndex + noSystems - 1);
                    writeToMultipleCells(startRange, endRange, "", 0, 1, 15, 21, true, false, true, System.Drawing.Color.Black, System.Drawing.Color.White, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignCenter, "Arial", 10);
                    //applyBorders(startRange, endRange);

                    startRange = new Location(l.getRow() + 1, colIndex);
                    endRange = new Location(l.getRow(), colIndex + noSystems - 1);

                    //applyGroups(startRange, endRange, TYPE_HEADING, i);
                }
                else if (l.getCellType() == TYPE_SUBHEADING)
                {
                    //Console.WriteLine("ROW: " + l.getRow() + " SUBHeading");
                    //Console.WriteLine("     Adding Data to " + colIndex);
                    Location startRange = new Location(l.getRow(), colIndex);
                    Location endRange = new Location(l.getRow(), colIndex + noSystems - 1);
                    writeToMultipleCells(startRange, endRange, "", 0, 1, 15, 21, true, false, true, System.Drawing.Color.DarkGray, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignTop, "Calibri", 11);

                    //applyBorders(startRange, endRange);
                    //applyGroups(startRange, endRange, TYPE_SUBHEADING, i);
                }
                else if (l.getCellType() == TYPE_SUBSUBHEADING2)
                {
                    //Console.WriteLine("ROW: " + l.getRow() + " SUBSUBHeading2");
                    //Console.WriteLine("     Adding Data to " + colIndex);
                    Location startRange = new Location(l.getRow(), colIndex);
                    Location endRange = new Location(l.getRow(), colIndex + noSystems - 1);
                    writeToMultipleCells(startRange, endRange, "", 0, 1, 15, 21, true, false, true, System.Drawing.Color.DarkGray, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignGeneral, Excel.XlVAlign.xlVAlignTop, "Calibri", 11);

                    //applyBorders(startRange, endRange);
                }
                else if (l.getCellType() == TYPE_SUBSUBHEADING)
                {
                    //Console.WriteLine("ROW: " + l.getRow() + " subsub");
                    //Console.WriteLine("     Adding Data to " + colIndex);
                    Location startRange = new Location(l.getRow(), colIndex);
                    Location endRange = new Location(l.getRow(), colIndex + noSystems - 1);
                    writeToMultipleCells(startRange, endRange, "", 0, 1, 15, 21, true, false, true, System.Drawing.Color.LightGray, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignGeneral, Excel.XlVAlign.xlVAlignTop, "Calibri", 11);

                    //applyBorders(startRange, endRange);
                }
                else if (l.getCellType() == TYPE_REQUIREMENT)
                {
                    //Console.WriteLine("ROW: " + l.getRow() + " req");
                    //Console.WriteLine("     Adding Data to " + colIndex);
                    Location startRange = new Location(l.getRow(), colIndex);
                    Location endRange = new Location(l.getRow(), colIndex + noSystems - 1);
                    writeToMultipleCells(startRange, endRange, "", 0, 1, 15, 21, true, false, true, System.Drawing.Color.AliceBlue, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignTop, "Calibri", 11);

                    applyBorders(startRange, endRange);
                }
                else if (l.getCellType() == TYPE_AVERAGE)
                {
                    //Console.WriteLine("ROW: " + l.getRow() + " Average");
                    //Console.WriteLine("     Adding Data to " + colIndex);
                    Location startRange = new Location(l.getRow(), colIndex);
                    Location endRange = new Location(l.getRow(), colIndex + noSystems - 1);
                    writeToMultipleCells(startRange, endRange, "", 0, 1, 15, 21, false, true, true, System.Drawing.Color.LightBlue, System.Drawing.Color.Black, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignBottom, "Arial", 10);

                    //applyBorders(startRange, endRange);
                }
                else if (l.getCellType() == TYPE_NOSCORE)
                {
                    //Console.WriteLine("ROW: " + l.getRow() + " NO SCORE");
                    //Console.WriteLine("     Adding Data to " + colIndex);
                    Location startRange = new Location(l.getRow(), colIndex);
                    Location endRange = new Location(l.getRow(), colIndex + noSystems - 1);
                    writeToMultipleCells(startRange, endRange, "", 0, 1, 15, 21, true, false, true, System.Drawing.Color.DeepSkyBlue, System.Drawing.Color.White, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignCenter, "Arial Narrow", 11);

                    applyBorders(startRange, endRange);
                }
                else if (l.getCellType() == TYPE_SYSTEM)
                {
                    //Console.WriteLine("ROW: " + l.getRow() + " System");
                    //Console.WriteLine("     Adding Data to " + colIndex);
                    Location startRange = new Location(l.getRow(), colIndex);
                    Location endRange = new Location(l.getRow(), colIndex + noSystems - 1);
                    writeToMultipleCells(startRange, endRange, "", 0, 8, 15, 55, true, false, true, System.Drawing.Color.DeepSkyBlue, System.Drawing.Color.White, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignTop, "Arial Narrow", 9);

                    applyBorders(startRange, endRange);
                }
                else if (l.getCellType() == TYPE_COMMENT)
                {
                    //Console.WriteLine("ROW: " + l.getRow() + " Comment");
                    //Console.WriteLine("     Adding Data to " + colIndex);
                    Location startRange = new Location(l.getRow(), colIndex);
                    Location endRange = new Location(l.getRow(), colIndex + noSystems - 1);
                    writeToMultipleCells(startRange, endRange, "", 0, 8, 15, 55, true, false, true, System.Drawing.Color.DeepSkyBlue, System.Drawing.Color.White, Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignTop, "Arial Narrow", 9);

                    applyBorders(startRange, endRange);
                }
                else
                {
                    //Console.WriteLine("ROW: " + l.getRow() + " Unknown");
                    //Console.WriteLine("     Adding Data to " + colIndex);
                    Location startRange = new Location(l.getRow(), colIndex);
                    Location endRange = new Location(l.getRow(), colIndex + noSystems - 1);
                    writeToMultipleCells(startRange, endRange, "", 0, 1, 15, 21, true, false, true, System.Drawing.Color.Black, System.Drawing.Color.White, Excel.XlHAlign.xlHAlignLeft, Excel.XlVAlign.xlVAlignCenter, "Arial", 10);

                    //applyBorders(startRange, endRange);
                }
                //}

            }
            Console.WriteLine("DONE");
            tasksCompleted++;



            // remove commented out region to perform a visual test that the virtual spreadsheet is the same as the one that is printed out
            /*for (int i = 0; i < virtualSpreadsheet.Count; i++)
            {
                Location l = virtualSpreadsheet[i];
                int row = l.getRow();
                Console.WriteLine(" Row " + row + " is type " + l.getCellType());
            }*/



            Console.Write("Applying Groups to SubSubHeadings...");

            // apply groupings to sub headings
            for (int i = 0; i < virtualSpreadsheet.Count; i++)
            {
                Location l = virtualSpreadsheet[i];
                if (l.getCellType() == TYPE_SUBSUBHEADING)
                {
                    Location startRange = new Location(l.getRow(), colIndex); // +1
                    Location endRange = new Location(l.getRow(), colIndex + noSystems - 1); // end row can be any parameter
                    applyGroups(startRange, endRange, TYPE_SUBSUBHEADING, i);
                }
                else if (l.getCellType() == TYPE_SUBHEADING) // also apply grouping to subheadings that contain requirements, and not subsubheadings
                {
                    Location startRange = new Location(l.getRow(), colIndex); // +1
                    Location endRange = new Location(l.getRow(), colIndex + noSystems - 1); // end row can be any parameter
                    applyGroups(startRange, endRange, TYPE_SUBHEADING, i);
                }
            }
            Console.WriteLine("DONE");
            tasksCompleted++;

            Console.Write("Applying Groups to SubHeadings...");

            // apply groupings to sub headings
            for (int i = 0; i < virtualSpreadsheet.Count; i++)
            {
                Location l = virtualSpreadsheet[i];
                if (l.getCellType() == TYPE_SUBHEADING)
                {
                    Location startRange = new Location(l.getRow(), colIndex); // +1
                    Location endRange = new Location(l.getRow(), colIndex + noSystems - 1); // end row can be any parameter
                    applyGroups(startRange, endRange, TYPE_SUBHEADING, i);
                }
            }
            Console.WriteLine("DONE");
            tasksCompleted++;

            Console.Write("Applying Groups to Headings...");

            // apply groupings to sub headings
            for (int i = 0; i < virtualSpreadsheet.Count; i++)
            {
                Location l = virtualSpreadsheet[i];
                if (l.getCellType() == TYPE_HEADING)
                {
                    Location startRange = new Location(l.getRow(), colIndex); // +1
                    Location endRange = new Location(l.getRow(), colIndex + noSystems - 1); // end row can be any parameter
                    applyGroups(startRange, endRange, TYPE_HEADING, i);
                }
            }
            Console.WriteLine("DONE");

            tasksCompleted++;

            Console.Write("Applying Number Formatting...");

            // then apply groupings to headings + apply hardcoded conditional formatting to requirements 
            for (int i = 0; i < virtualSpreadsheet.Count; i++)
            {
                Location l = virtualSpreadsheet[i];
                if (l.getCellType() == TYPE_HEADING)
                {
                    Location startRange = new Location(l.getRow() + 1, colIndex);
                    Location endRange = new Location(0, colIndex + noSystems - 1); // note: row is negligable here because it gets changed anyways
                    //applyGroups(startRange, endRange, TYPE_HEADING, i);
                }
                if (l.getCellType() == TYPE_SUBHEADING || l.getCellType() == TYPE_HEADING)
                {
                    Location startRange = new Location(l.getRow(), colIndex);
                    Location endRange = new Location(l.getRow(), colIndex + noSystems - 1);
                    applyNumberFormatting(startRange, endRange, l.getCellType());
                }
                if (l.getCellType() == TYPE_REQUIREMENT)
                {
                    Location startRange = new Location(l.getRow(), colIndex);
                    Location endRange = new Location(l.getRow(), colIndex + noSystems - 1);
                    //applyConditionalFormatting(startRange, endRange);
                    applyNumberFormatting(startRange, endRange, l.getCellType());
                }
                if (l.getCellType() == TYPE_AVERAGE)
                {
                    Location startRange = new Location(l.getRow(), colIndex);
                    Location endRange = new Location(l.getRow(), colIndex + noSystems - 1);
                    applyNumberFormatting(startRange, endRange, l.getCellType());
                }
            }
            Console.WriteLine("DONE");
            tasksCompleted++;
            sheetSuccessfullySetUp = true;
            return 0;

        }

        private void makeTextFile()
        {
            string[] arr = new string[virtualSpreadsheet.Count];

            for (int i = 0; i < arr.Length; i++)
                arr[i] = virtualSpreadsheet[i].getCellType().ToString();

            System.IO.File.WriteAllLines(@"C:\Users\Alan\Desktop\virtualSpread Sheet", arr);
        }



        public int findNoSystems()
        {
            int currentRow = startRow;
            int currentCol = startCol;
            int noSystems = 0;
            Boolean endOfDocumentFound = false;

            while (!endOfDocumentFound)
            {
                if (MySystems.GetValue(currentRow, currentCol) == null)
                    endOfDocumentFound = true;
                else
                {
                    String val = MySystems.GetValue(currentRow, currentCol).ToString();
                    if (val != "")
                    {
                        systemNames.Add(val);
                        //Console.WriteLine("Found " + val);
                        noSystems++;
                    }

                }
                currentRow++;
            }
            return noSystems;
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




        private void writeToSingleCell(Location cellLocation, string value, int mode)
        {
            int row = cellLocation.getRow();
            int col = cellLocation.getColumn();


            if (mode == 1 || mode == 2)
                Console.WriteLine("Writing to " + col + "  (" + cellLocation.getExcelAddress() + "): " + value);
            if (mode == 0 || mode == 2)
                newTableSheet.Cells[row, col] = value;
        }

        private void writeToSingleCell(Location cellLocation, string value, int mode, int cellType, double cellWidth, double cellHeight, bool boldFlag, bool italicsFlag, bool wrapTextFlag, System.Drawing.Color cellColour,
            System.Drawing.Color textColour, Excel.XlHAlign horizAlignment, Excel.XlVAlign vertAlignment, String fontName, int fontSize)
        {
            int row = cellLocation.getRow();
            int col = cellLocation.getColumn();

            if (cellType != TYPE_IGNORE && !alreadyWrittenTo.Contains(row))
            {
                virtualSpreadsheet.Add(new Location(row, col, cellType));
                alreadyWrittenTo.Add(row);
            }

            if (mode == 1 || mode == 2)
                Console.WriteLine("Writing to " + col + "  (" + cellLocation.getExcelAddress() + "): " + value);
            if (mode == 0 || mode == 2)
            {
                newTableSheet.Cells[row, col] = value;
                Excel.Range r = newTableSheet.Cells[row, col] as Excel.Range;
                newTableSheet.Columns[col].ColumnWidth = cellWidth;
                newTableSheet.Rows[row].RowHeight = cellHeight;
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

            if (cellType != TYPE_IGNORE && !alreadyWrittenTo.Contains(startRow))
            {
                virtualSpreadsheet.Add(new Location(startRow, startCol, cellType));
                alreadyWrittenTo.Add(startRow);
            }

            if (mode == 1 || mode == 2)
                Console.WriteLine("Writing to " + startCol + "  (" + startRange + ":" + endRange + ") " + value);
            if (mode == 0 || mode == 2)
            {


                newTableSheet.Columns[startCol].ColumnWidth = cellWidth;
                newTableSheet.Rows[startRow].RowHeight = cellHeight;
                Excel.Range r = newTableSheet.get_Range(startRange, endRange); //Excel.Range r = newTableSheet.Cells[startRow, startCol] as Excel.Range;
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

        private void applyBorders(Location startRangeLocation, Location endRangeLocation)
        {
            int startRow = startRangeLocation.getRow();
            int startCol = startRangeLocation.getColumn();
            int endRow = endRangeLocation.getRow();
            int endCol = endRangeLocation.getColumn();

            String startRange = getExcelColumnName(startCol) + startRow.ToString();
            String endRange = getExcelColumnName(endCol) + endRow.ToString();
            //Console.WriteLine("Applying border to range " + startRange + ":" + endRange);
            Excel.Range r = newTableSheet.get_Range(startRange, endRange); //Cells[row, col] as Excel.Range;
            r.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
            r.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
            r.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
            r.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
            r.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
            r.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
        }

        private void applyGroups(Location startRangeLocation, Location endRangeLocation, int type, int index)
        {
            index++;                          // intially increment index becuase we want to start grouping on next row down on the startlocation row
            int emptyHeadingCount = 0;        // counter used to identify if a heading, subheading or subsubheading is empty, i.e contains no information beneath it, if this counter remains 0, we don't group that heading, subheading or subsubheading

            Boolean endOfGroupFound = false; // boolean to exit loop when the end range location of the group has been determined

            while (!endOfGroupFound)
            {
                if (index >= (virtualSpreadsheet.Count - 1))
                {
                    endOfGroupFound = true;
                    index = virtualSpreadsheet.Count;
                }
                else if (virtualSpreadsheet[index].getCellType() <= type)
                    endOfGroupFound = true;
                else
                {
                    index++;
                    emptyHeadingCount++;
                }
            }

            int row = virtualSpreadsheet[index - 1].getRow();

            endRangeLocation.setRow(row);
            startRangeLocation.setColumn(startRangeLocation.getColumn() - 2);
            startRangeLocation.setRow(startRangeLocation.getRow() + 1);

            if (emptyHeadingCount != 0)      // only apply grouping if not 0
            {
                Console.WriteLine("     Applying group to " + startRangeLocation.getExcelAddress() + ":" + endRangeLocation.getExcelAddress());
                Excel.Range r = newTableSheet.get_Range(startRangeLocation.getExcelAddress(), endRangeLocation.getExcelAddress()); //Cells[row, col] as Excel.Range;
                r.Rows.Group();
            }


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

        public int getNewStartRow()
        {
            return newStartRow;
        }

        public int getNewStartColumn()
        {
            return newStartCol;
        }


        public void conditionalFormat()
        {
            int colIndex = newStartCol + 2;
            for (int i = 0; i < virtualSpreadsheet.Count; i++)
            {
                Location l = virtualSpreadsheet[i];
                if (l.getCellType() == TYPE_REQUIREMENT)
                {
                    Location startRange = new Location(l.getRow(), colIndex);
                    Location endRange = new Location(l.getRow(), colIndex + noSystems - 1);
                    applyConditionalFormatting(startRange, endRange);
                }
            }
        }

        private void applyConditionalFormatting(Location startRangeLocation, Location endRangeLocation)
        {
            int startRow = startRangeLocation.getRow();
            int startCol = startRangeLocation.getColumn();
            int endRow = endRangeLocation.getRow();
            int endCol = endRangeLocation.getColumn();

            String startRange = getExcelColumnName(startCol) + startRow.ToString();
            String endRange = getExcelColumnName(endCol) + endRow.ToString();
            //Console.WriteLine("Applying Conditional Formatting to range " + startRange + ":" + endRange);
            Excel.Range r = newTableSheet.get_Range(startRange, endRange);

            //r.FormatConditions.AddColorScale(8711167).Value = 50;
            //r.FormatConditions.AddColorScale(65535).Value = 100;

            Excel.ColorScale cfColorScale = (Excel.ColorScale)(r.FormatConditions.AddColorScale(2));

            // Set the minimum threshold to red (0x000000FF) and maximum threshold
            // to blue (0x00FF0000).

            Int32 red = 0x000000FF;
            Int32 yellow = 0x000ff0FF;

            cfColorScale.ColorScaleCriteria[1].FormatColor.Color = red;
            cfColorScale.ColorScaleCriteria[2].FormatColor.Color = yellow;
            cfColorScale.ColorScaleCriteria[1].Type = Excel.XlConditionValueTypes.xlConditionValueNumber;
            cfColorScale.ColorScaleCriteria[2].Type = Excel.XlConditionValueTypes.xlConditionValueNumber;
            cfColorScale.ColorScaleCriteria[1].Value = 0;
            cfColorScale.ColorScaleCriteria[2].Value = 10;

        }

        private void applyNumberFormatting(Location startRangeLocation, Location endRangeLocation, int cellType)
        {
            int startRow = startRangeLocation.getRow();
            int startCol = startRangeLocation.getColumn();
            int endRow = endRangeLocation.getRow();
            int endCol = endRangeLocation.getColumn();

            String startRange = getExcelColumnName(startCol) + startRow.ToString();
            String endRange = getExcelColumnName(endCol) + endRow.ToString();
            //Console.WriteLine("Applying Number Formatting to range " + startRange + ":" + endRange);
            Excel.Range r = newTableSheet.get_Range(startRange, endRange);
            if (cellType == TYPE_REQUIREMENT)
                r.NumberFormat = "0.0";// "###,##.#";
            if (cellType == TYPE_HEADING || cellType == TYPE_SUBHEADING)
            {
                r.Style = "Percent";// "###,##.#%";
                //r.NumberFormat = "0.00%";

            }
            if (cellType == TYPE_AVERAGE)
                r.NumberFormat = "0.00";
        }


    }



}
