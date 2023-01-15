using CsvHelper.Configuration.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderWriter.Classes_Json
{
    internal class ClientesCSV
    {
        [Name("SITE")] public string Site { get; set; }
        [Name("NOME")] public string Nome { get; set; }
        [Name("POTENCIA")] public string Potencia { get; set; }
    }
}