using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;

namespace ShowBestPlayer
{
    public class Program
    {
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        private static Excel.Range XlRange = null;
        public int SkipRow;
        public int AccumPlayerTotals;

        static void Main(string[] args)
        {
            string StartTime = DateTime.Now.ToString();
            Console.WriteLine("Simple List - Start time: {0}", StartTime);
            string Player = null;
            string Position = null;
            double PassYards = 0;
            double RushYards = 0;
            int Index;
 
            SetUpExcel();  //Open EXCEL and spreadsheet

            int LastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            List<AccumPlayerTotals> PlayerTotalList = new List<AccumPlayerTotals>();

            for (int i = 2; i < LastRow; i++)   // Start at row 2 to bypass column headers
            {
                Position = XlRange.Cells.Value2[i, 22].ToString();
                Player = XlRange.Cells.Value2[i, 2].ToString();

                if ((Player != " " && Player != null) && (Position == "QB" || Position == "RB"))  // Bypass rows without valid player, position
                {
                    Index = i;
                    ProcessPlayer(Player, Position, PassYards, RushYards, Index, PlayerTotalList);
                }
            }

            GetLeaders(PlayerTotalList, out IOrderedEnumerable<AccumPlayerTotals> QBLeader, out IOrderedEnumerable<AccumPlayerTotals> RBLeader);
            
            //-- close the collection
            CloseSheet();
        }

        //######################################################################
        private static void SetUpExcel()
        {
            MyApp = new Excel.Application
            {
                Visible = false  //suppress the display of the Excel spreadsheet
            };
            string XLS_PATH = "C:\\Users\\Dennis.Langmeyer\\Desktop\\NFL_Small_Set.xlsx";
            MyBook = MyApp.Workbooks.Open(XLS_PATH);
            MySheet = (Excel.Worksheet)MyBook.Sheets["Sheet1"];
            XlRange = MySheet.UsedRange;
        }

        //######################################################################
        private static void ProcessPlayer(string Player, string Position, double PassYards, double RushYards, int Index, List<AccumPlayerTotals> PlayerTotalList)
        {
            VerifyYards(out PassYards, out RushYards, Index);

            if (PlayerTotalList.Count == 0)  // List is empty - Insert the first entry 
            {
                AddEntry(Player, Position, PassYards, RushYards, PlayerTotalList);
            }
            else  // the list has existing entries
            {
                Index = PlayerTotalList.FindIndex(x => x.Player.Contains(Player));
                if (Index == -1) // - Player not found in list - insert
                {
                    AddEntry(Player, Position, PassYards, RushYards, PlayerTotalList);
                }
                else // Player found in list - accummulate totals
                {
                    PlayerTotalList[Index].PassYards += PassYards;
                    PlayerTotalList[Index].RushYards += RushYards;
                }
            }
        }

        //######################################################################
        private static void GetLeaders(List<AccumPlayerTotals> PlayerTotalList, out IOrderedEnumerable<AccumPlayerTotals> QBLeader, out IOrderedEnumerable<AccumPlayerTotals> RBLeader)
        {
            //---Sort in descending order to put highest yard total as top row
            QBLeader = PlayerTotalList.OrderByDescending(x => x.PassYards);
            RBLeader = PlayerTotalList.OrderByDescending(x => x.RushYards);
            WriteStats(QBLeader, RBLeader);
        }

        //######################################################################
        private static void WriteStats(IOrderedEnumerable<AccumPlayerTotals> QBLeader, IOrderedEnumerable<AccumPlayerTotals> RBLeader)
        {
            //-- Write the solution 
            Console.WriteLine("Passing   Leader " + QBLeader.First().Player + " " + QBLeader.First().PassYards);
            Console.WriteLine("Rushing   Leader " + RBLeader.First().Player + " " + RBLeader.First().RushYards);
            string endTime = DateTime.Now.ToString();
            Console.WriteLine("Simple List - End time: {0}", endTime);
            Console.ReadKey();
        }

        //######################################################################
        private static void CloseSheet()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(XlRange);
            Marshal.ReleaseComObject(MySheet);
            
            MyBook.Close();     //-- close the Excel spreadsheet
            Marshal.ReleaseComObject(MyBook);
            
            MyApp.Quit();       //-- close Excel 
            Marshal.ReleaseComObject(MyApp);
        }

        //######################################################################
        private static void AddEntry(string player, string position, double PassYards, double RushYards, List<AccumPlayerTotals> PlayerTotalList)
        {
            AccumPlayerTotals AccumPlayerTotals = new AccumPlayerTotals
            {
                Player = player,
                Position = position,
                PassYards = PassYards,
                RushYards = RushYards
            };
            PlayerTotalList.Add(AccumPlayerTotals);
        }

        //######################################################################
        private static void VerifyYards(out double PassYards, out double RushYards, int i)
        {
            string CheckNull = " ";
            CheckNull = Convert.ToString(XlRange.Cells[i, 11].Value2);
            if (string.IsNullOrWhiteSpace(CheckNull))
            {
                PassYards = 0;
            }
            else
            {
                PassYards = XlRange.Cells.Value2[i, 11];
            }
            CheckNull = Convert.ToString(XlRange.Cells[i, 15].Value2);
            if (string.IsNullOrWhiteSpace(CheckNull))
            {
                RushYards = 0;
            }
            else
            {
                RushYards = XlRange.Cells.Value2[i, 15];
            }

            CheckNull = Convert.ToString(XlRange.Cells[i, 19].Value2);
        }
    }
}
