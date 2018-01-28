using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Vml.Spreadsheet;

namespace ExcelValidator
{
    public class ReadExcel
    {
        private string _fileName;
        private int _bullSize = 50;

        public ReadExcel(string fileName)
        {
            this._fileName = fileName;
        }

        public List<User> ReadFile()
        {
            string uploadPath = HttpContext.Current.Server.MapPath("~/App_Data");
            var xlWorkbook = new XLWorkbook(uploadPath + "/" + this._fileName);
            var workSheet = xlWorkbook.Worksheet("Users");

            int firstRow = 2;
            int lastRow = workSheet.LastRowUsed().RowNumber();
            int firstColumn = 1;
            int lastColumn = workSheet.LastColumnUsed().ColumnNumber();
            int numberOfRecords = lastRow - firstRow;


            var range = workSheet.Range(firstRow,firstColumn,lastRow,lastColumn);

            var xlTableRows = range.AsTable().DataRange.Rows().ToList();

            List<User> users = new List<User>();

            foreach (var row in xlTableRows)
            {
                User user = new User();
                var cells = row.Cells().ToList();
                user.Email = cells[0].Value.ToString();
                user.Username = cells[1].Value.ToString();
                user.Password = cells[2].Value.ToString();
                user.Firstname = cells[3].Value.ToString();
                user.Lastname = cells[4].Value.ToString();
                user.UserType = cells[5].Value.ToString();
                user.AlternativeEmail = cells[6].Value.ToString();
                user.AccountCode = cells[7].Value.ToString();
                user.PhoneNumber = cells[8].Value.ToString();
                user.CostCenterName = cells[9].Value.ToString();

                users.Add(user);
            }

            return users;
        }
    }
}
