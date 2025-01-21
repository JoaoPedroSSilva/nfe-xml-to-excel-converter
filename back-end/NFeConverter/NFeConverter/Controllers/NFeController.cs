using Microsoft.AspNetCore.Mvc;
using System.Xml;
using System.Data;
using ClosedXML.Excel;

namespace NFeConverter.Controllers {
    [ApiController]
    [Route("api/[controller]")]
    public class NFeController : Controller {
        [HttpPost("convert")]
        public IActionResult ConvertXmlToExcel(IFormFile file) {
            if (file == null || file.Length == 0) {
                return BadRequest("Arquivo inválido.");
            }

            if (!file.FileName.EndsWith(".xml")) {
                return BadRequest("Formato de arquivo inválido. " +
                    "Envie um arquivo XML.");
            }

            try {
                var doc = new XmlDocument();
                using (var stream = file.OpenReadStream()) {
                    doc.Load(stream);
                }

                var dataTable = new DataTable();
                dataTable.Columns.Add("Chave");
                dataTable.Columns.Add("Valor Total");
                dataTable.Columns.Add("Emitente");
                dataTable.Columns.Add("Destinatário");
                dataTable.Columns.Add("Produto");
                dataTable.Columns.Add("Quantidade");

                Console.WriteLine(doc.SelectNodes("//nfeProc").Count);

                foreach (XmlNode node in doc.SelectNodes("//nfeProc")) {
                    var row = dataTable.NewRow();
                    row["Chave"] = node.Attributes["Id"]?.Value.Replace("NFe", ""); 
                    row["Valor Total"] = node.SelectSingleNode("total/ICMSTot/vNF")?.InnerText;
                    row["Emitente"] = node.SelectSingleNode("emit/xNome")?.InnerText;
                    row["Destinatário"] = node.SelectSingleNode("dest/xNome")?.InnerText;

                    var produto = node.SelectSingleNode("det/prod");
                    row["Produto"] = produto?.SelectSingleNode("xProd")?.InnerText;
                    row["Quantidade"] = produto?.SelectSingleNode("qCom")?.InnerText;

                    dataTable.Rows.Add(row);
                }

                using (var workbook = new XLWorkbook()) {
                    var worksheet = workbook.Worksheets.Add(dataTable, "NFe");

                    worksheet.Cell(1, 1).Value = "Chave";
                    worksheet.Cell(1, 2).Value = "Valor";
                    
                    for (int i = 0; i < dataTable.Rows.Count; i++) {
                        worksheet.Cell(i + 2, 1).Value = dataTable.Rows[i]["Chave"]?.ToString();
                        worksheet.Cell(i + 2, 1).Value = dataTable.Rows[i]["Valor"]?.ToString();
                            }

                    worksheet.Columns().AdjustToContents();

                    using (var stream = new System.IO.MemoryStream()) {
                        workbook.SaveAs(stream);
                        stream.Seek(0, System.IO.SeekOrigin.Begin);
                        return File(stream.ToArray(),
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            "Converted.xlsx");
                    }
                }
            } catch (Exception ex) {
                return StatusCode(500, $"Erro interno: {ex.Message}");
            }
        }
    }
}
