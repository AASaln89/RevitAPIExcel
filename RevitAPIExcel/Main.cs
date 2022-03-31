using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPIExcel
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            string wallsInfo = string.Empty;

            var walls = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Walls)
                .Cast<Wall>()
                .ToList();

            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Walls.xlsx");

            using (FileStream stream = new FileStream (excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("sheet1");

                int rowIndex = 0;
                foreach(var wall in walls)
                {
                    sheet.SetCellValue(rowIndex, columnIndex:0, wall.Name);
                    sheet.SetCellValue(rowIndex, columnIndex:0, wall.LookupParameter("Volume").AsString());
                    rowIndex++;
                }
                workbook.Write(stream);
                workbook.Close();
            }

            System.Diagnostics.Process.Start(excelPath);
            return Result.Succeeded;
        }
    }
}
