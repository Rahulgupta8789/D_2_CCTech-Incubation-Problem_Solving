/*
2.Marks obtained by students in 5 different subjects in a class are to be processed to generate
  the following results :
  Total marks, Percentage, Grade for each student,
  Any special performance class topper, subject topper results;
*/

using System;
using Microsoft.Office.Interop.Excel;    //Add Library use Excel sheet

namespace HelloWorld.ReadFile
{
    class ExcelSheet
    {
        string path = "";
        _Application excel = new Application();
        Workbook wb;                              //Excel file 
        Worksheet ws;                             //Excel sheet store

        public ExcelSheet(string path, int sheet) //constructor
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        public string ReadCell(int r, int c)  //Read each cell row & column in Excel sheet 
        {
            r++; c++;

            if (ws.Cells[r, c].Value2 != null)
                return ws.Cells[r, c].Value2.ToString();  
            else
                return "";
        }

        public void PrintResult(int rollNo)        //Result print basis of Roll No
        {
            int rowNo = FindRowNo(rollNo);         

            if (rowNo != 0)          
            {
                int totalMarks = TotalMarks(rowNo);   //call totalMarks
                float percentage = GetPercent(rowNo); //call percentage
                string grade = GetGrade(rowNo);       //call grade

                //Display Result
                Console.WriteLine($"Roll No : {rollNo}\t\tName    : {ReadCell(rowNo, 1)}");
                Console.WriteLine("**********************************************************************************");
                Console.WriteLine($"{ReadCell(0, 2)}\t\t{ReadCell(0, 3)}\t\t{ReadCell(0, 4)}\t{ReadCell(0, 5)}\t\t{ReadCell(0, 6)}");
                Console.WriteLine("**********************************************************************************");
                Console.WriteLine($"{ReadCell(rowNo, 2)}\t\t{ReadCell(rowNo, 3)}\t\t{ReadCell(rowNo, 4)}\t\t{ReadCell(rowNo, 5)}\t\t{ReadCell(rowNo, 6)}");
                Console.WriteLine($"\nTotal Marks : {totalMarks}");
                Console.WriteLine($"\nPercentage : {percentage}%\t\tGrade : {grade}");

            }
            else
            {
                Console.WriteLine("Roll no not found!!!");
            }
        }

        //==========================================================================
        
        public void PrintAllStudentResult()
        {
            int rowNo = 2;
            Console.WriteLine("\n Total marks, Percentage, Grade for each student Results :\n");
            Console.Write($"Roll No \tName \t Math\t English\tComputer\tScience\t " +
                 $"Social Science\t\tTotal Marks\tPercentage\tGrade\n");
            Console.WriteLine("******************************************************************" +
                "*******************************************************************\n");

            while (ReadCell(rowNo, 0) != "")
            {
                int totalMarks = TotalMarks(rowNo);
                float percentage = GetPercent(rowNo);
                string grade = GetGrade(rowNo);
                Console.Write($"{ReadCell(rowNo, 0)}\t\t{ReadCell(rowNo, 1)}\t " +
                    $"{ReadCell(rowNo, 2)}\t  {ReadCell(rowNo, 3)}\t\t {ReadCell(rowNo, 4)}" +
                    $"\t\t  {ReadCell(rowNo, 5)}\t\t{ReadCell(rowNo, 6)}\t\t");
                Console.Write($"{totalMarks}\t");
                Console.Write($"\t{percentage}%\t\t{grade}\n\n");

                rowNo++;
            }



            Console.WriteLine("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" +
                 "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX\n");
            Console.WriteLine("\nClass Topper Result :\n");
            PrintResult(GetTopper());



            Console.WriteLine("\nXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" +
                 "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX\n");
            Console.WriteLine("\nSubject Topper Result :\n");
            PrintSubjectTopper();
        }
        //============================================================================
        public int FindRowNo(int rollNo)
        {
            int row = 2;

            while (ReadCell(row, 0) != "")
            {
                if (int.Parse(ReadCell(row, 0)) == rollNo)
                {
                    return row;
                }
                row++;
            }
            return 0;
        }

        public int TotalMarks(int row)      //Method TotalMarks
        {
            int totalMarks = 0;

            for (int i = 2; i <= 6; i++)
            {
                totalMarks += int.Parse(ReadCell(row, i));
            }

            return totalMarks;
        }

        public float GetPercent(int rowNO)   //Method GetPercent
        {
            float percentage = (float)TotalMarks(rowNO) / int.Parse(ReadCell(rowNO, 7)) * 100;
            return percentage;
        }
        public string GetGrade(int rowNo)
        {
            float percentage = GetPercent(rowNo);
            bool isPass = IsPass(rowNo);



            if (percentage >= 90 && isPass) return "A+";
            else if (percentage >= 80 && isPass) return "A";
            else if (percentage >= 70 && isPass) return "B+";
            else if (percentage >= 60 && isPass) return "B";
            else if (percentage >= 50 && isPass) return "C";
            else if (percentage >= 40 && isPass) return "D";
            else if (percentage >= 35 && isPass) return "E";
            else return "Fail";
        }
        public bool IsPass(int row)
        {
            int marks = 0;
            bool isPass = true;

            for (int i = 2; i <= 6; i++)
            {
                marks = int.Parse(ReadCell(row, i));
                if (marks < 35)
                {
                    isPass = false;
                    break;
                }
            }
            return isPass;
        }
        public int GetTopper()          //Method GetTopper
        {
            int row = 2;
            int highestMarks = 0;
            int totalMarks = 0;
            int topperRoll = 0;

            while (ReadCell(row, 0) != "")      //It reads every line until it the end line is null in excel sheet.
            {
                totalMarks = TotalMarks(row);
                if (totalMarks > highestMarks)
                {
                    highestMarks = totalMarks;
                    topperRoll = int.Parse(ReadCell(row, 0));
                }
                row++;
            }
            return topperRoll;
        }

        
        // Returns Array of Roll no of Subject topper 
        public int[] SubjectTopper()     //Method SubjectTopper
        {
            int row = 2;
            int[] subjctTopperList = new int[5];
            int st1 = 0, st2 = 0, st3 = 0, st4 = 0, st5 = 0;    //To store subject topper temporarily 

            while (ReadCell(row, 0) != "")
            {
                if (st1 < int.Parse(ReadCell(row, 2)))
                {
                    st1 = int.Parse(ReadCell(row, 2));
                    subjctTopperList[0] = int.Parse(ReadCell(row, 0));
                }

                if (st2 < int.Parse(ReadCell(row, 3)))
                {
                    st2 = int.Parse(ReadCell(row, 3));
                    subjctTopperList[1] = int.Parse(ReadCell(row, 0));
                }

                if (st3 < int.Parse(ReadCell(row, 4)))
                {
                    st3 = int.Parse(ReadCell(row, 4));
                    subjctTopperList[2] = int.Parse(ReadCell(row, 0));
                }

                if (st4 < int.Parse(ReadCell(row, 5)))
                {
                    st4 = int.Parse(ReadCell(row, 5));
                    subjctTopperList[3] = int.Parse(ReadCell(row, 0));
                }

                if (st5 < int.Parse(ReadCell(row, 6)))
                {
                    st5 = int.Parse(ReadCell(row, 6));
                    subjctTopperList[4] = int.Parse(ReadCell(row, 0));
                }

                row++;

            }
            return subjctTopperList;
        }

        public void PrintSubjectTopper()     //Method PrintSubjectTopper
        {
            int[] subjectTopperList = SubjectTopper();
            int rowNo = 0;
            for (int i = 0; i <= 4; i++)
            {
                rowNo = FindRowNo(subjectTopperList[i]);
                Console.WriteLine($"Name : {ReadCell(rowNo, 1)}\t\tSubject : {ReadCell(0, i + 2)}\t\tMarks : {ReadCell(rowNo, i + 2)}");
            }

        }

    }
    class StudentResult
    {
        static void Main(string[] args) 
        {
            //ExcelSheet e1 = new ExcelSheet("C:\\Users\\Admin\\Downloads\\Problem_Solving\\Problem_Solving\\ProblemSolving - Copy (3)", 1);
            //ExcelSheet e1 = new ExcelSheet("C:\\Users\\Admin\\source\\repos\\Problem_Solving\\Problem_Solving\\ProblemSolving - Copy (4)", 1);
            ExcelSheet e1 = new ExcelSheet("C:\\Users\\Admin\\source\\repos\\Problem_Solving\\Problem_Solving\\ProblemSolving - Copy (3)", 1);
            
            
            ////e1.PrintResult(3);
            
            e1.PrintAllStudentResult();

            //e1.PrintResult(e1.GetTopper());

            //e1.PrintSubjectTopper();
        }
    }
}

