using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using excel_generator.Models;
using Microsoft.AspNetCore.Mvc;

namespace excel_generator.Controllers
{
    [ApiController]
    public class MainController : ControllerBase
    {
        private readonly StoreContext _db;
        public MainController(StoreContext db)
        {
            _db = db;
        }

        [HttpGet("excel/generate")]
        public FileContentResult GenerateExcelFile()
        {
            List<Student> students = new List<Student>();
            for (int i = 1; i < 9; i++)
            {
                students.Add(new Student()
                {
                    Id = i,
                    Name = "Student " + i,
                    StudentId = 100 + i
                });
            }

            using (var workbook = new XLWorkbook())
            {
                // Tab
                var worksheet = workbook.Worksheets.Add("Student");
                var currentRow = 1;

                //Header
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "Name";
                worksheet.Cell(currentRow, 3).Value = "StudentId";

                //Body
                foreach (var student in students)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = student.Id;
                    worksheet.Cell(currentRow, 2).Value = student.Name;
                    worksheet.Cell(currentRow, 3).Value = student.StudentId;
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return File(
                        content, "application/vnd.openxmlformat-officedocument.spreadsheetml.sheet",
                        "StudentData.xlsx"
                    );
                }
            }
        }

        [HttpGet("excel/generate-from-db")]
        public FileContentResult GenerateExcelFileFromDB()
        {
            List<Student> students = _db.Students.ToList();
            
            using (var workbook = new XLWorkbook())
            {
                // Tab
                var worksheet = workbook.Worksheets.Add("Student");
                var currentRow = 1;

                //Header
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "Name";
                worksheet.Cell(currentRow, 3).Value = "StudentId";

                //Body
                foreach (var student in students)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = student.Id;
                    worksheet.Cell(currentRow, 2).Value = student.Name;
                    worksheet.Cell(currentRow, 3).Value = student.StudentId;
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return File(
                        content, "application/vnd.openxmlformat-officedocument.spreadsheetml.sheet",
                        "StudentDataFromDB.xlsx"
                    );
                }
            }
        }

    }

}