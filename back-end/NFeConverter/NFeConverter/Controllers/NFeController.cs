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
            if (!ValidateXmlFile(file, out var validationMessage)) {
                return BadRequest(validationMessage);
            }

            try {
                var doc = LoadXmlDocument(file);
                var dataTable = CreateDataTable();
                PopulateDataTable(doc, dataTable);
                return GenerateExcelFile(dataTable);
            } catch (Exception ex) {
                return StatusCode(500, $"Erro interno: {ex.Message}");
            }
        }

        private IActionResult GenerateExcelFile(DataTable dataTable) {
            using (var workbook = new XLWorkbook()) {
                var worksheet = workbook.Worksheets.Add(dataTable, "NFe");
                worksheet.Columns().AdjustToContents();

                using (var stream = new System.IO.MemoryStream()) {
                    workbook.SaveAs(stream);
                    stream.Seek(0, System.IO.SeekOrigin.Begin);
                    return File(stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "Converted.xlsx");
                }
            }
        }

        private void PopulateDataTable(XmlDocument doc, DataTable dataTable) {

            var row = dataTable.NewRow();
            row["Chave"] = doc.GetElementsByTagName("infNFe")[0]?.Attributes["Id"]?.Value.Replace("NFe", "");
            row["Valor Total"] = doc.GetElementsByTagName("vNF")[0]?.InnerText;
            row["Emitente"] = doc.GetElementsByTagName("xNome")[0]?.InnerText;
            row["Destinatário"] = doc.GetElementsByTagName("xNome")[1]?.InnerText;
            row["Produto"] = doc.GetElementsByTagName("xProd")[0]?.InnerText;
            row["Quantidade"] = doc.GetElementsByTagName("qCom")[0]?.InnerText;

            dataTable.Rows.Add(row);

        }

        private DataTable CreateDataTable() {
            var dataTable = new DataTable();
            dataTable.Columns.Add("Chave");
            dataTable.Columns.Add("Valor Total");
            dataTable.Columns.Add("Emitente");
            dataTable.Columns.Add("Destinatário");
            dataTable.Columns.Add("Produto");
            dataTable.Columns.Add("Quantidade");
            return dataTable;
        }

        private bool ValidateXmlFile(IFormFile file, out string validationMessage) {
            if (file == null || file.Length == 0) {
                validationMessage = "Arquivo inválido.";
                return false;
            }

            if (!file.FileName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) {
                validationMessage = "Formato de arquivo inválido. Envie um arquivo XML NF-e válido.";
                return false;
            }

            validationMessage = null;
            return true;
        }

        private XmlDocument LoadXmlDocument(IFormFile file) {
            var doc = new XmlDocument();
            using (var stream = file.OpenReadStream()) {
                doc.Load(stream);
            }
            return doc;
        }
    }
}
