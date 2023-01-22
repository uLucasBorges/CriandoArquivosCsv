// See https://aka.ms/new-console-template for more information
using CriandoArquivosCsv.Models;
using CriandoArquivosCsv.Services;


var c = new ExcelService();
var funcionarios = new List<Funcionario>();
funcionarios.Add(new Funcionario(100, "Bianca do vapo"));
//funcionarios.Add(new Funcionario(2, "Teste"));
//funcionarios.Add(new Funcionario(3, "aaaa"));
//funcionarios.Add(new Funcionario(4, "bbb"));

c.Insert(funcionarios);
//c.ObterExcel();


