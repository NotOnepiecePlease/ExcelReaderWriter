using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReaderWriter.Classes;

namespace ExcelReaderWriter.Classes_Json
{
    public class Site
    {
        public string nome { get; set; }
        public List<Cliente> clientes { get; set; }
    }
}