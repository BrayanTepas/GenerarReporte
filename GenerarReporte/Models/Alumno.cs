using System;
using System.Collections.Generic;

namespace GenerarReporte.Models;

public partial class Alumno
{
    public int Id { get; set; }

    public string? Nombre { get; set; }

    public int? Edad { get; set; }
}
