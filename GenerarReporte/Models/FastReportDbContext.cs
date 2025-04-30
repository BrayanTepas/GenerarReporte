using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;

namespace GenerarReporte.Models;

public partial class FastReportDbContext : DbContext
{
    public FastReportDbContext()
    {
    }

    public FastReportDbContext(DbContextOptions<FastReportDbContext> options)
        : base(options)
    {
    }

    public virtual DbSet<Alumno> Alumnos { get; set; }

    
    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Alumno>(entity =>
        {
            entity.HasKey(e => e.Id).HasName("PK__Alumnos__3214EC073B31D82F");

            entity.Property(e => e.Nombre).HasMaxLength(100);
        });

        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
