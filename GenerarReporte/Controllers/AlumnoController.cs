using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using GenerarReporte.Models;
using ClosedXML.Excel;
using System.Data;
using QuestPDF.Fluent;
using Xceed.Words.NET;

namespace GenerarReporte.Controllers
{
    public class AlumnoController : Controller
    {
        private readonly FastReportDbContext _context;

        public AlumnoController(FastReportDbContext context)
        {
            _context = context;
        }

        // GET: Alumno
        public async Task<IActionResult> Index()
        {
            return View(await _context.Alumnos.ToListAsync());
        }

        // GET: Alumno/Details/5
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var alumno = await _context.Alumnos
                .FirstOrDefaultAsync(m => m.Id == id);
            if (alumno == null)
            {
                return NotFound();
            }

            return View(alumno);
        }

        // GET: Alumno/Create
        public IActionResult Create()
        {
            return View();
        }

        // POST: Alumno/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("Id,Nombre,Edad")] Alumno alumno)
        {
            if (ModelState.IsValid)
            {
                _context.Add(alumno);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            return View(alumno);
        }

        // GET: Alumno/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var alumno = await _context.Alumnos.FindAsync(id);
            if (alumno == null)
            {
                return NotFound();
            }
            return View(alumno);
        }

        // POST: Alumno/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("Id,Nombre,Edad")] Alumno alumno)
        {
            if (id != alumno.Id)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(alumno);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!AlumnoExists(alumno.Id))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            return View(alumno);
        }

        // GET: Alumno/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var alumno = await _context.Alumnos
                .FirstOrDefaultAsync(m => m.Id == id);
            if (alumno == null)
            {
                return NotFound();
            }

            return View(alumno);
        }

        // POST: Alumno/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var alumno = await _context.Alumnos.FindAsync(id);
            if (alumno != null)
            {
                _context.Alumnos.Remove(alumno);
            }

            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        private bool AlumnoExists(int id)
        {
            return _context.Alumnos.Any(e => e.Id == id);
        }

        // Método asincrónico que genera y exporta un archivo Excel con la lista de alumnos
        public async Task<IActionResult> ExportarExcel() /*int id*/
        {
            // Hacemos la consulta a la base de datos y la guardamos en una variable
            var _alumnos = await _context.Alumnos.ToListAsync();

            // Le damos nombre al archivo Excel que se va a generar
            var nombreArchivo = $"ListadoAlumnos.xlsx";

            // Creamos un DataTable para colocar los registros
            // Se requiere el using: System.Data;
            // DataTable representa una tabla en memoria
            DataTable dataTable = new DataTable("Lista1");

            // Agregamos las columnas que tendrá el DataTable
            dataTable.Columns.AddRange(new DataColumn[]
            {
        new DataColumn("Nombre "),
        new DataColumn("Edad"),
            });

            // Colocamos el resultado de la consulta dentro del DataTable
            foreach (var alumno in _alumnos)
            {
                // Agregamos una fila por cada alumno con su nombre y edad
                dataTable.Rows.Add(alumno.Nombre, alumno.Edad);
            }

            // Usamos la biblioteca ClosedXML para generar el archivo Excel
            // Se requiere el using: ClosedXML.Excel;
            using (XLWorkbook wb = new XLWorkbook())
            {
                // Agregamos el DataTable como una hoja del libro de Excel
                wb.Worksheets.Add(dataTable);

                // Creamos un stream de memoria para guardar el archivo generado
                using (MemoryStream strem = new MemoryStream())
                {
                    // Guardamos el archivo Excel en el stream
                    wb.SaveAs(strem);

                    // Retornamos el archivo Excel como una respuesta de descarga
                    return File(strem.ToArray(), "Aplicaction/vnd.openxmlformats.spreadsheetml.sheet",
                            nombreArchivo);
                }
            }

            // En caso de no retornar el archivo, se redirige a la acción Index (esta línea nunca se ejecutará realmente)
            return RedirectToAction(nameof(Index));
        }

        // Método asincrónico que exporta un PDF con la lista de alumnos
        public async Task<IActionResult> ExportarPdf()
        {
            // Obtiene la lista de alumnos de la base de datos de forma asincrónica
            var alumnos = await _context.Alumnos.ToListAsync();

            // Crea un flujo de memoria para almacenar el PDF generado
            var stream = new MemoryStream();

            // Crea el documento PDF utilizando la librería QuestPDF
            Document.Create(container =>
            {
                // Define una página dentro del documento
                container.Page(page =>
                {
                    // Establece el margen de la página
                    page.Margin(30);

                    // Define el contenido de la página como una columna
                    page.Content().Column(col =>
                    {
                        // Agrega un título en negrita con tamaño de fuente 14
                        col.Item().Text("Nombre - Edad").Bold().FontSize(14);

                        // Itera sobre la lista de alumnos para mostrarlos en el PDF
                        foreach (var alumno in alumnos)
                        {
                            // Agrega una línea por cada alumno con su nombre y edad
                            col.Item().Text($"{alumno.Nombre} - {alumno.Edad}").FontSize(12);
                        }
                    });
                });
            }).GeneratePdf(stream); // Genera el PDF y lo escribe en el flujo de memoria

            // Reinicia la posición del flujo al inicio
            stream.Position = 0;

            // Devuelve el archivo PDF generado como una respuesta para descarga
            return File(stream.ToArray(), "application/pdf", "ListadoAlumnos.pdf");
        }

        // Método que exporta un documento Word con la lista de alumnos
        public IActionResult ExportarWord()
        {
            // Obtiene la lista de alumnos desde la base de datos
            var alumnos = _context.Alumnos.ToList();

            // Crea un flujo de memoria para almacenar el documento Word
            using (var stream = new MemoryStream())
            {
                // Crea un nuevo documento Word utilizando la librería DocX
                using (var doc = DocX.Create(stream))
                {
                    // Inserta un título en el documento con formato en negrita y tamaño 14
                    var titulo = doc.InsertParagraph("Nombre - Edad")
                        .Bold()
                        .FontSize(14)
                        .SpacingAfter(15); // Espaciado después del título

                    // Itera sobre la lista de alumnos para agregarlos al documento
                    foreach (var alumno in alumnos)
                    {
                        // Inserta una línea con el nombre y edad de cada alumno
                        doc.InsertParagraph($"{alumno.Nombre} - {alumno.Edad}");
                    }

                    // Guarda el documento en el flujo de memoria
                    doc.Save();
                }

                // Establece la posición del flujo al inicio para que pueda ser leído
                stream.Seek(0, SeekOrigin.Begin);

                // Devuelve el archivo Word generado como respuesta para descarga
                return File(stream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document", // Tipo MIME para Word
                    "alumnos.docx"); // Nombre del archivo descargado
            }
        }
    }
}
