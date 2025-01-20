using Microsoft.AspNetCore.Mvc;
using System.Xml;
using System.Data;
using ClosedXML.Excel;

namespace NFeConverter.Controllers {
    [ApiController]
    [Route("api/[controller]")]
    public class NFeController : Controller {
        [HttpPost("convert")]
        public IActionResult ConvertXmlToExcel([FromForm] IFormFile file) {
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
                dataTable.Columns.Add("ExemploColuna1");
                dataTable.Columns.Add("ExemploColuna2");

                foreach (XmlNode node in doc.SelectNodes("//SeuXpathAqui")) {
                    var row = dataTable.NewRow();
                    row["ExemploColuna1"] = node["Elemento1"]?.InnerText;
                    row["ExemploColuna2"] = node["Elemento2"]?.InnerText;
                    dataTable.Rows.Add(row);
                }

                using (var workbook = new XLWorkbook()) {
                    var worksheet = workbook.Worksheets.Add(dataTable, "Dados");
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
