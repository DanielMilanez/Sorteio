using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace Sorteio {
    public class Pessoa {
        public string Name { get; set; }
        public string Number { get; set; }

        public static List<Pessoa> LeitorXls() {
            var resposta = new List<Pessoa>();

            FileInfo ArquivoExistente = new FileInfo(fileName: "C:\\Teste.xlsx");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(ArquivoExistente)) {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[PositionID: 0];
                int colCount = worksheet.Dimension.End.Column;

                int rowCount = worksheet.Dimension.End.Row;

                for (int row = 2; row <= rowCount; row++) {
                    var pessoa = new Pessoa();

                    pessoa.Name = worksheet.Cells[row, Col: 2].Value.ToString();
                    pessoa.Number = worksheet.Cells[row, Col: 1].Value.ToString();

                    resposta.Add(pessoa);
                }
            }
            return resposta;
        }

    }
}
