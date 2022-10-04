﻿using Microsoft.AspNetCore.Mvc;
using PdfExport.Models;
using System.Collections.Generic;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.IO;

namespace PdfExport.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class PdfFileController : ControllerBase
    {
        [HttpGet("GetEmpLists")]

        public List<EmpListModel> GetEmpLists()
        {
            List<EmpListModel> EmpData = new List<EmpListModel>
            {
                new EmpListModel{slNo=1,EmpName="Tushar",dept="DEv",Desg="SE"},
                new EmpListModel{slNo=2,EmpName="Atul",dept="deploy",Desg="LE"},
                new EmpListModel{slNo=3,EmpName="Shivanand",dept="test",Desg="SE"}

            };

            return EmpData;
        }

        [HttpPost("GeneratePDF")]

        public async Task<ActionResult> GenratePdf()
        {
            using (var workbook = new XLWorkbook())
            {
                var workSheet = workbook.Worksheets.Add("Employees");
                var currentRow = 1;

                workSheet.Cell(currentRow, 1).Value = "slNo";
                workSheet.Cell(currentRow, 2).Value = "EmpName";
                workSheet.Cell(currentRow, 3).Value = "Dept";
                workSheet.Cell(currentRow, 4).Value = "Desg";

                var data = GetEmpLists();
                foreach (var i in data)
                {
                    currentRow++;
                    workSheet.Cell(currentRow, 1).Value = i.slNo;
                    workSheet.Cell(currentRow, 2).Value = i.EmpName;
                    workSheet.Cell(currentRow, 3).Value = i.dept;
                    workSheet.Cell(currentRow, 4).Value = i.Desg;
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return File(
                                  content,
                                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                  "Students.xlsx"
                               );
                }
            }

        }
    }
}
