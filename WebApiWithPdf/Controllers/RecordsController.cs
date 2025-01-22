using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using WebApiWithPdf.Data;
using WebApiWithPdf.Models;

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

        // GET: api/records/download
        [HttpGet("download")]
        public async Task<IActionResult> DownloadRecordsAsPdf()
        {
            var records = await _context.Records.ToListAsync();

            using (var memoryStream = new MemoryStream())
            {
                var writer = new PdfWriter(memoryStream);
                var pdf = new PdfDocument(writer);
                var document = new Document(pdf);

                // Add Title
                document.Add(new Paragraph("User Records List                                    Page no:1")


                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetBold()
                    .SetFontSize(18)
                    .SetMarginBottom(20));



                // Create Table with Column Headers
                var table = new Table(UnitValue.CreatePercentArray(new float[] { 1, 3, 3, 1, 3 }))
                    .UseAllAvailableWidth()
                    .SetMarginTop(10);

                table.AddHeaderCell(new Cell().Add(new Paragraph("ID").SetBold()));
                table.AddHeaderCell(new Cell().Add(new Paragraph("Name").SetBold()));
                table.AddHeaderCell(new Cell().Add(new Paragraph("Surname").SetBold()));
                table.AddHeaderCell(new Cell().Add(new Paragraph("Age").SetBold()));
                table.AddHeaderCell(new Cell().Add(new Paragraph("Phone Number").SetBold()));

                // Populate Table with Records
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
    }
}
