using BestFreightProject.Dtos;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BestFreightProject.Services
{
    public interface IExcelService
    {
        FileStream GetExcel(string path);
        void CreateExcelFile(ExcelCreateDto excelCreate);

    }
}
