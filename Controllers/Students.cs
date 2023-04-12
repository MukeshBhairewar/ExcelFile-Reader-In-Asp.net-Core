using ExcelDataReader;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using ReadExcelFile.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ReadExcelFile.Controllers
{
    public class Students : Controller
    {
        [HttpGet]
        public IActionResult Index(List<Student> students=null)
        {
            students = students == null ? new List<Student>() : students;
            return View(students);
        }

        [HttpPost]
        [Obsolete]
        public IActionResult Index(IFormFile file,[FromServices] IHostingEnvironment hostingEnivronment)
        {
            string filename = $"{hostingEnivronment.WebRootPath}\\ExcelFile\\{file.FileName}";
           using(FileStream fileStream= System.IO.File.Create(filename))
            {
                file.CopyTo(fileStream);
                fileStream.Flush();
            }
            var students = this.GetStudentList(file.FileName);
            return Index(students);

        }

        private List<Student> GetStudentList(string fName)
        {
            List<Student> students = new List<Student>();
            var fileName = $"{Directory.GetCurrentDirectory()}{@"\wwwroot\ExcelFile"}" + "\\" + fName;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using(var stream = System.IO.File.Open(fileName,FileMode.Open, FileAccess.Read))
            {
                using(var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while(reader.Read())
                    {

                        if (reader.Depth == 0) // check if it's the first row
                        {
                            continue; // skip the first row
                        }

                        students.Add(new Student()
                            {
                                Id = reader.GetValue(0).ToString(),
                                Name = reader.GetValue(1).ToString(),
                                Age = reader.GetValue(2).ToString(),
                                Email = reader.GetValue(3).ToString(),
                                MobileNo = reader.GetValue(4).ToString()

                            });
                        
                    }
                }
            }
            return students;
        }
    }
}
