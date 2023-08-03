using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Ink;

namespace Sorteio {
    public class Program {
        static void Main(string[] args) {

            int min = 0;
            Console.Write("Informe o número maximo para ser sorteado: ");
            int max = int.Parse(Console.ReadLine());

            Random random = new Random(); //criando objeto que gera números "aleátorios"
            int numeroSorteado = random.Next(min, max);


            var pessoas = Pessoa.LeitorXls();

            foreach (var itens in pessoas) {
                if (numeroSorteado.ToString() == itens.Number) {
                    Console.WriteLine(value: $"Número: {itens.Number}\nNome: {itens.Name}");
                    Console.WriteLine();
                }
            }
        }
    }
}
