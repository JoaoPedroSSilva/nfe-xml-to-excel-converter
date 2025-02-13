using Microsoft.AspNetCore.Mvc;
using System.Xml;
using System.Data;
using ClosedXML.Excel;

namespace NFeConverter.Controllers {
    [ApiController]
    [Route("api/[controller]")]
    public class NFeController : Controller {
        [HttpPost("convert")]
        public IActionResult ConvertXmlToExcel([FromForm] List<IFormFile> files) {
            if (files == null || files.Count == 0) {
                return BadRequest("Nenhum arquivo enviado.");
            }

            var dataTable = CreateDataTable();

            foreach (var file in files) {
                string validation = ValidateXmlFile(file);
                if (validation != "") {
                    return BadRequest(validation);
                }

                try {
                    var doc = LoadXmlDocument(file);
                    PopulateDataTable(doc, dataTable);
                } catch (Exception ex) {
                    return StatusCode(500, $"Erro ao processar o arquivo '{file.FileName}': {ex.Message}");
                }
            }

            try {
                return GenerateExcelFile(dataTable);
            } catch (Exception ex) {
                return StatusCode(500, $"Erro ao gerar o arquivo Excel: {ex.Message}");
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

            var dateEmi = doc.GetElementsByTagName("dhEmi")[0]?.InnerText;
            if (!string.IsNullOrEmpty(dateEmi)) {
                var shortDate = dateEmi.Substring(0, dateEmi.IndexOf('T'));

                if (DateTime.TryParse(shortDate, out var formatDate)) {
                    row["Data Emissão"] = formatDate.ToString("dd/MM/yyyy");
                } else {
                    row["Data Emissão"] = "Data inválida";
                }
            } else {
                row["Data Emissão"] = "Data inválida";
            }

            row["Nota"] = doc.GetElementsByTagName("nNF")[0]?.InnerText;
            row["Valor Total"] = doc.GetElementsByTagName("vNF")[0]?.InnerText.Replace(".", ",");
            row["Emitente"] = doc.GetElementsByTagName("xNome")[0]?.InnerText;
            row["CNPJ Emitente"] = doc.GetElementsByTagName("CNPJ")[0]?.InnerText;
            row["Destinatário"] = doc.GetElementsByTagName("xNome")[1]?.InnerText;
            row["CNPJ Destinatário"] = doc.GetElementsByTagName("CNPJ")[1]?.InnerText;
            row["Natureza Operação"] = doc.GetElementsByTagName("natOp")[0]?.InnerText;
            row["Chave"] = doc.GetElementsByTagName("infNFe")[0]?.Attributes["Id"]?.Value.Replace("NFe", "");

            dataTable.Rows.Add(row);
        }

        private DataTable CreateDataTable() {
            var dataTable = new DataTable();
            dataTable.Columns.Add("Data Emissão");
            dataTable.Columns.Add("Nota");
            dataTable.Columns.Add("Valor Total");
            dataTable.Columns.Add("Emitente");
            dataTable.Columns.Add("CNPJ Emitente");
            dataTable.Columns.Add("Destinatário");
            dataTable.Columns.Add("CNPJ Destinatário");
            dataTable.Columns.Add("Natureza Operação");
            dataTable.Columns.Add("Chave");
            return dataTable;
        }

        private string ValidateXmlFile(IFormFile file) {
            string validationMessage = "";
            if (file == null || file.Length == 0) {
                validationMessage = "Arquivo inválido.";
                return validationMessage;
            }

            if (!file.FileName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) {
                validationMessage = "Formato de arquivo inválido. Envie um arquivo XML NF-e válido.";
                return validationMessage;
            }

            return validationMessage;
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
