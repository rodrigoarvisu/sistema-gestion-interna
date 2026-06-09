CREATE DATABASE PROGRAMA_PC_BOMBEROS 

USE PROGRAMA_PC_BOMBEROS 

CREATE TABLE Usuarios(
id_usuario INT IDENTITY(1, 1) PRIMARY KEY,
nombre_usuario VARCHAR (50) NOT NULL UNIQUE,
contrasena VARCHAR (255) NOT NULL,
nombre_completo VARCHAR (100) NOT NULL,
rol VARCHAR (20) CHECK (rol IN('admin', 'usuario')) NOT NULL,
estatus VARCHAR(20) CHECK (estatus IN ('Activo','Inactivo')) NOT NULL,
fecha_creacion DATETIME DEFAULT GETDATE(),
fecha_modificacion DATETIME NULL );

INSERT INTO Usuarios
(nombre_usuario, contrasena, nombre_completo, rol, estatus)
VALUES 
('prueba', '03ac674216f3e15c761ee1a5e255f067953623c8b388b4459e13f978d7c846f4', 'PRUEBA', 'admin', 'Activo');



SELECT * FROM Usuarios
DELETE FROM Usuarios

CREATE TABLE Capacitacion(
id_capacitacion INT IDENTITY(1,1) PRIMARY KEY,
nombre_inmueble VARCHAR(100) NOT NULL,
tipo VARCHAR(50) NOT NULL, 
direccion VARCHAR (150),
fecha DATE,

--CAPACITACIONES
introduccion_proteccion_civil BIT, 
primeros_aux BIT, 
combate_incendios BIT,
plan_emergencia BIT,
evacuacion_inmuebles BIT,
manejo_sustancias_quimicas BIT,
busqueda_rescate BIT,
practica_combate_contra_incendios BIT, 
simulacro BIT,

--CONFORMACIÓN 
uipc BIT,
comite BIT,

--SIMULACRO
hipotesis VARCHAR(100));

ALTER TABLE Capacitacion
ADD estatus VARCHAR(50),
poblacion_fija INT,
poblacion_flotante INT,
total INT, 
hora_inicio TIME, 
hora_fin TIME,
tiempo_evacuacion VARCHAR (50),
alertamiento BIT, 
observaciones VARCHAR (500);

ALTER TABLE Capacitacion
ADD simulacro_gabinete VARCHAR (50);

ALTER TABLE Capacitacion
ADD usuario_creacion VARCHAR(50),
fecha_creacion DATETIME, 
usuario_modificacion VARCHAR (50),
fecha_modificacion DATETIME;

SELECT * FROM Capacitacion;

TRUNCATE TABLE Capacitacion

SELECT fecha_modificacion FROM Capacitacion

SELECT id_capacitacion, usuario_creacion, usuario_modificacion
FROM Capacitacion
ORDER BY id_capacitacion DESC


CREATE TABLE Difusiones(
id_difusion INT IDENTITY (1, 1) PRIMARY KEY, 
red_social VARCHAR(50),
tipo_publicacion VARCHAR(50),
fecha_publicacion DATE,
enlace VARCHAR(500),
responsable VARCHAR(100),
descripcion VARCHAR(1000),
diseño_contenido VARCHAR(50),
usuario_creacion VARCHAR(50),
fecha_creacion DATETIME, 
usuario_modificacion VARCHAR(50),
fecha_modificacion DATETIME);

SELECT * FROM Difusiones 


CREATE TABLE Vinculacion (
id_vinculacion INT IDENTITY(1, 1) PRIMARY KEY,
sector VARCHAR(100),
nombre VARCHAR(150),
direccion VARCHAR(500),
fecha DATE,
hora_creacion TIME, 
hora_termino TIME,
coordinador VARCHAR(150),
jefe_comunicacion VARCHAR(150),
jefe_evaluacion VARCHAR(150),
jefe_primeros_auxilios VARCHAR(150),
jefe_incendios VARCHAR(150));

ALTER TABLE Vinculacion
ADD usuario_creacion VARCHAR(50),
fecha_creacion DATETIME, 
usuario_modificacion VARCHAR(50),
fecha_modificacion DATETIME;


CREATE TABLE Integrantes(
id_integrante INT IDENTITY(1, 1) PRIMARY KEY,
id_vinculacion INT NOT NULL, 
cargo VARCHAR(100),
nombre VARCHAR(150),
telefono VARCHAR(80),
correo VARCHAR(200),

FOREIGN KEY (id_vinculacion)
REFERENCES vinculacion(id_vinculacion)
ON DELETE CASCADE);

ALTER TABLE Vinculacion
ADD observaciones NVARCHAR(MAX);


SELECT * FROM Vinculacion

SELECT * FROM Integrantes

SELECT * FROM Integrantes WHERE id_vinculacion = 1;
SELECT * FROM Integrantes WHERE id_vinculacion = 2;

TRUNCATE TABLE Capacitacion
TRUNCATE TABLE Difusiones

DELETE FROM Vinculacion
DBCC CHECKIDENT ('Vinculacion', RESEED, 0);

TRUNCATE TABLE Integrantes