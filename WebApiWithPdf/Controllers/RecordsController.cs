using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using WebApiWithPdf.Data;
using WebApiWithPdf.Models;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Kernel.Events;
using iText.Kernel.Pdf.Canvas;

namespace WebApiWithPdf.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class RecordsController : ControllerBase
    {
        private readonly ApplicationDbContext _context;

        public RecordsController(ApplicationDbContext context)
        {
            _context = context;
        }

        // POST: api/records
        [HttpPost]
        public async Task<IActionResult> CreateRecord([FromBody] Record record)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            await _context.Records.AddAsync(record);
            await _context.SaveChangesAsync();
            return CreatedAtAction(nameof(GetRecordById), new { id = record.Id }, record);
        }

        // GET: api/records/{id}
        [HttpGet("{id}")]
        public async Task<IActionResult> GetRecordById(int id)
        {
            var record = await _context.Records.FindAsync(id);
            if (record == null)
            {
                return NotFound();
            }
            return Ok(record);
        }

        // Download records as PDF
        [HttpGet("download/pdf")]
        public async Task<IActionResult> DownloadRecordsAsPdf()
        {
            var records = await _context.Records.ToListAsync();

            using (var memoryStream = new MemoryStream())
            {
                var writer = new PdfWriter(memoryStream);
                var pdf = new PdfDocument(writer);
                var document = new Document(pdf);

                pdf.AddEventHandler(PdfDocumentEvent.END_PAGE, new PageNumberEventHandler(document));

                var currentDateTime = DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt");

                document.Add(new Paragraph("User Records List")
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetBold()
                    .SetFontSize(18)
                    .SetMarginBottom(5));

                // No longer adding the date/time at the top of the document, as we will add it below the page number

                var table = new Table(UnitValue.CreatePercentArray(new float[] { 1, 3, 3, 1, 3 }))
                    .UseAllAvailableWidth()
                    .SetMarginTop(10);

                table.AddHeaderCell(new Cell().Add(new Paragraph("ID").SetBold()));
                table.AddHeaderCell(new Cell().Add(new Paragraph("Name").SetBold()));
                table.AddHeaderCell(new Cell().Add(new Paragraph("Surname").SetBold()));
                table.AddHeaderCell(new Cell().Add(new Paragraph("Age").SetBold()));
                table.AddHeaderCell(new Cell().Add(new Paragraph("Phone Number").SetBold()));

                foreach (var record in records)
                {
                    table.AddCell(new Paragraph(record.Id.ToString()));
                    table.AddCell(new Paragraph(record.Name));
                    table.AddCell(new Paragraph(record.Surname));
                    table.AddCell(new Paragraph(record.Age.ToString()));
                    table.AddCell(new Paragraph(record.PhoneNumber));
                }

                document.Add(table);

                document.Close();

                var pdfBytes = memoryStream.ToArray();
                return File(pdfBytes, "application/pdf", "Records.pdf");
            }
        }

        // Download records as Excel
        [HttpGet("download/excel")]
        public async Task<IActionResult> DownloadRecordsAsExcel()
        {
            var records = await _context.Records.ToListAsync();

            using (var memoryStream = new MemoryStream())
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.AddWorksheet("User Records");

                    // Add header row
                    worksheet.Cell(1, 1).Value = "ID";
                    worksheet.Cell(1, 2).Value = "Name";
                    worksheet.Cell(1, 3).Value = "Surname";
                    worksheet.Cell(1, 4).Value = "Age";
                    worksheet.Cell(1, 5).Value = "Phone Number";

                    // Populate data
                    int row = 2;
                    foreach (var record in records)
                    {
                        worksheet.Cell(row, 1).Value = record.Id;
                        worksheet.Cell(row, 2).Value = record.Name;
                        worksheet.Cell(row, 3).Value = record.Surname;
                        worksheet.Cell(row, 4).Value = record.Age;
                        worksheet.Cell(row, 5).Value = record.PhoneNumber;
                        row++;
                    }

                    // Save the file to the memory stream
                    workbook.SaveAs(memoryStream);
                }

                var excelBytes = memoryStream.ToArray();
                return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Records.xlsx");
            }
        }

        // Event Handler for Page Numbers
        private class PageNumberEventHandler : IEventHandler
        {
            private readonly Document _document;

            public PageNumberEventHandler(Document document)
            {
                _document = document;
            }

            public void HandleEvent(Event @event)
            {
                var pdfEvent = (PdfDocumentEvent)@event;
                var page = pdfEvent.GetPage();
                var pdfCanvas = new PdfCanvas(page.NewContentStreamAfter(), page.GetResources(), pdfEvent.GetDocument());
                var pageSize = page.GetPageSize();

                var canvas = new Canvas(pdfCanvas, pageSize);

                int pageNumber = pdfEvent.GetDocument().GetPageNumber(page);

                // Display page number at the top-right corner
                canvas
                    .ShowTextAligned($"Page no: {pageNumber}",
                        pageSize.GetRight() - 40,
                        pageSize.GetTop() - 20,
                        TextAlignment.RIGHT)
                    .Close();

                // Display the current date and time below the page number
                canvas
                    .ShowTextAligned($"Date: {DateTime.Now:MM/dd/yyyy hh:mm:ss tt}",
                        pageSize.GetRight() - 40,
                        pageSize.GetTop() - 40,
                        TextAlignment.RIGHT)
                    .Close();
            }
        }
    }
}
