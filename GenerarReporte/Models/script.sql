CREATE DATABASE FastReportDb
USE FastReportDb;

CREATE TABLE Alumnos (
    Id INT PRIMARY KEY IDENTITY,
    Nombre NVARCHAR(100),
    Edad INT
);

INSERT INTO Alumnos (Nombre, Edad) VALUES ('Carlos', 20), ('Lucía', 22), ('Ana', 25);