using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CriandoArquivosCsv.Models
{
    public class Funcionario
    {
        public int Id { get; set; }
        public string Name { get; set; }

        public Funcionario(int id, string name)
        {
            Id = id;
            Name = name;
        }
    }
}
