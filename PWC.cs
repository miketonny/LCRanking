using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace LCRanking
{
    public partial class PWC
    {
        private void PWC_Load(object sender, RibbonUIEventArgs e)
        {
          
        }

        private void btnRun_Click(object sender, RibbonControlEventArgs e)
        {
            var wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            List<LCItem> items = new List<LCItem>();
            // open up csv file 
            var file = @"C:\Users\micha\Desktop\LC.csv";
            if (!File.Exists(file)) return;
            var reader = new StreamReader(File.OpenRead(file));
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                var values = line.Split(',');
                // add items
                items.Add(new LCItem { Name = values[0], Point = Convert.ToInt32(values[1]), ItemId = string.IsNullOrEmpty(values[2].Trim()) ? 0 : Convert.ToInt32(values[2].Trim()), Deducted = Convert.ToDouble(values[4]), Priority = 0, PassPoints = Convert.ToInt32(values[5])});
            }
            reader.Close();

            // parses items into sheet 
            Worksheet rankingSheet = ((Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);
            if (rankingSheet.Cells[2, 1].Text != "Kel'Thuzad") return;
            // get attendances
            Worksheet attendanceSheet = ((Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[2]);
            if (attendanceSheet.Cells[4, 1].Text != "Arhat") return;
            List<Attendance> attendanceList = GetAttendance(attendanceSheet);
            // get point deduction
            Worksheet deductionSheet = ((Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[3]);
            List<Deduction> deductList = GetDeduction(deductionSheet);
            // get point passes
            Worksheet passesSheet = ((Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[4]);
            List<Passes> passesList = GetPasses(passesSheet);
            items.ForEach(p => p.Priority = CalculatePriority(p.Point, attendanceList.Where(x => x.Name.Equals(p.Name)).FirstOrDefault()?.AttendPerc, deductList.Where(x => x.Name.Equals(p.Name)).ToList(), passesList.Where(x=> x.Name.Equals(p.Name) && x.ItemId == p.ItemId).FirstOrDefault()?.Point )); // calc priorities for each person and item..  
            var bossItems = rankingSheet.Range["C3", "C190"];
            foreach (Range itm in bossItems)
            {
                // get all players with this item
                if (string.IsNullOrEmpty(itm.Text)) continue;
                
                var playersWanted = items.Where(x => x.ItemId == Convert.ToInt32(itm.Text)).ToList(); 
                if (playersWanted.Count() == 0) continue; // move next
                // found players want this item, rank them..
                int colIndx = 4; // start with "D" column
                playersWanted.OrderByDescending(x => x.Priority).ToList().ForEach(p =>
                {
                    if (!attendanceList.Any(x => x.Name.Equals(p.Name))) return;
                    rankingSheet.Cells[itm.Row, colIndx].Value = $"{p.Name} : {p.Priority}";
                    if (colIndx < 9) rankingSheet.Cells[itm.Row, colIndx].Style = GetCellStyle(colIndx);
                    colIndx++;
                });
            }
            // find the right workbook, start working..
 
        }

        private List<Passes> GetPasses(Worksheet passesSheet)
        {
            List<Passes> passList = new List<Passes>();
            var names = passesSheet.Range["A2", "A999"];
            int row = 2; // record start from row 2
            foreach (Range name in names)
            {
                // get all players with this item
                if (string.IsNullOrEmpty(name.Text)) continue;
                passList.Add(new Passes { Name = name.Text, Point = Convert.ToDouble(passesSheet.Range[$"C{row}"].Value2), ItemId = Convert.ToInt32(passesSheet.Range[$"B{row}"].Value2) });
                row++;
            }
            return passList;
        }

        private List<Deduction> GetDeduction(Worksheet deductionSheet)
        {
            List<Deduction> atts = new List<Deduction>();
            var names = deductionSheet.Range["A2", "A99"];
            int row = 2; // record start from row 2
            foreach (Range name in names)
            {
                // get all players with this item
                if (string.IsNullOrEmpty(name.Text)) continue;
                atts.Add(new Deduction { Name = name.Text, DeductPoint = Convert.ToDouble(deductionSheet.Range[$"B{row}"].Value2) });
                row++;
            }
            return atts;

        }

        private List<Attendance> GetAttendance(Worksheet attendanceSheet)
        {
            List<Attendance> atts = new List<Attendance>();
            var names = attendanceSheet.Range["A4", "A50"]; // active players only
            int row = 4; // record start from row 4
            foreach (Range name in names)
            {
                // get all players with this item
                if (string.IsNullOrEmpty(name.Text)) continue;
                atts.Add(new Attendance { Name = name.Text, AttendPerc = Convert.ToInt32(attendanceSheet.Range[$"U{row}"].Value2) });
                row++;
            }
            return atts;
        }

        private dynamic GetCellStyle(int colIndx)
        {
            var color = "Normal";
            switch (colIndx)
            {
                case 4:
                    color = "No1";
                    break;
                case 5:
                    color = "No5";
                    break;
                case 6:
                    color = "No2";
                    break;
                case 7:
                    color = "No3";
                    break;
                case 8:
                    color = "No4";
                    break;
                default:
                    break;
            }
            return color; 
        }

        // item point + passed times * 0.4 + attendance * 0.1 - deducted = Priority for item.
        private double CalculatePriority(int pt, int? attnd, List<Deduction> deductList, double? passes)
        {
            double deduct = 0;
            if (deductList.Count() == 0) deduct = 0;
            else
            {
                switch (pt)
                {
                    case 50:
                        deduct = deductList[0].DeductPoint;
                        break;
                    case 49:
                        if (deductList.Count() > 1) deduct = deductList[1].DeductPoint;
                        break;
                    case 48:
                        if (deductList.Count() > 2) deduct = deductList[2].DeductPoint;
                        break;
                    case 47:
                        if (deductList.Count() > 3) deduct = deductList[3].DeductPoint;
                        break;
                    default:
                        break;
                }
            }

            return (Convert.ToInt32(pt) + (passes ?? 0) + (Convert.ToInt32(attnd ?? 0) * 0.1) + deduct);
        }
    }
}
