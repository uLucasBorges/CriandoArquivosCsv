using System;
using System.Data;
using System.Data.OleDb;
using ClosedXML.Excel;
using CriandoArquivosCsv.Models;

namespace CriandoArquivosCsv.Services
{
    public class ExcelService
    {
        //nome desejado para a planilha
        private static String name = "Resultados";
        //caminho de onde vai ficar a planilha
        private static String _Arquivo = @$"C:\Users\lucas\Downloads\{name}.xlsx";
        private String _StringConexao = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0 Xml;HDR=YES;ReadOnly=False';", _Arquivo);
        private OleDbConnection _olecon;



        public void Insert(List<Funcionario> funcionarios)
        {
            if (!this.ValidarExistencia())
            {
                this.CriarPlanilha();
            }
            
            
            foreach(var item in funcionarios)
            {
                _olecon = new OleDbConnection(_StringConexao);
                _olecon.Open();

               
          
                OleDbCommand _oleCmd = new OleDbCommand();
                _oleCmd.Connection = _olecon;

                string _Consulta;
                var CodFunci = item.Id;
                var NomeFunci = item.Name;

                _Consulta = "INSERT INTO [Resultados$] ";
                _Consulta += "([CodFunci],[NomeFunci]) ";
                _Consulta += "VALUES ";
                _Consulta += "(@CodFunci,@NomeFunci)";

                _oleCmd.CommandText = _Consulta;
                _oleCmd.Parameters.Add("@CodFunci", OleDbType.Integer).Value = Convert.ToInt32(item.Id);
                _oleCmd.Parameters.Add("@NomeFunci", OleDbType.VarChar, 255).Value = item.Name;
                _oleCmd.ExecuteNonQuery();

               
                _oleCmd.Parameters.Clear();

                var xls = new XLWorkbook(_Arquivo);
                xls.Save();


                Console.WriteLine("Dados Incluídos...");
            }





            Console.WriteLine("Dados Inclusos.");
        }

        private void Obter()
        {


            var xls = new XLWorkbook(_Arquivo);
            var planilha = xls.Worksheets.First(w => w.Name == "Resultados");
            var totalLinhas = planilha.Rows().Count();


            for (int i = 1; i <= totalLinhas; i++)
            {
                Console.WriteLine($"\nLINHA {i}\n");
                Console.WriteLine("ID: " + planilha.Cell("A" + i).Value.ToString());
                Console.WriteLine("Nome: " + planilha.Cell("B" + i).Value.ToString());
            }

        }


        private bool ValidarExistencia()
        {
            if (File.Exists(_Arquivo))
                return true;

            CriarPlanilha();
            return true;
        }

        private bool CriarPlanilha()
        {
            using (var workbook = new XLWorkbook())
            {
                var planilha = workbook.Worksheets.Add(name);

                workbook.SaveAs(_Arquivo);

                return true;

            }
        }

    }
}
