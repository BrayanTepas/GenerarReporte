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

        public async Task<IActionResult> ExportarExcel()
        {
            //Hacemos la consulta y la gardamos 
            var _alumnos = await _context.Alumnos.ToListAsync();

            // Le damos nombre al archivo que se va generar 
            var nombreArchivo = $"ListadoAlumnos.xlsx";

            //hacemos uso de DataTable para colocar los registros 
            //hacemos uso de los using - System.Data; 
            //                         - DataTable = System.Data.DataTable; 

            DataTable dataTable = new DataTable("Lista1");
            dataTable.Columns.AddRange(new DataColumn[]
            {
                new DataColumn("Nombre "),
                new DataColumn("Edad"),
            });

            //Ccolocamos el resultado de la consulta al dataTable 
            foreach (var alumno in _alumnos)
            {
                dataTable.Rows.Add(alumno.Nombre, alumno.Edad);
            }
            //Analista de sistemas  

            //Agregamos using XLWorkbook para poder generar el archivo ..  
            //ClosedXML.Excel; 
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dataTable);
                using (MemoryStream strem = new MemoryStream())
                {
                    wb.SaveAs(strem);
                    return File(strem.ToArray(), "Aplicaction/vnd.openxmlformats.spreadsheetml.sheet",
                            nombreArchivo);
                }
            }
            ;
            return RedirectToAction(nameof(Index));
        }
    }
}
