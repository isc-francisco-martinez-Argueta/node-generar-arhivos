const exceljs = require("exceljs");
// Crea un archivo excel
const Workbook = new exceljs.Workbook();
// Crea hojas en el archivo excel
const hojaincidencias = Workbook.addWorksheet("incidencias");
const hojapersonasDB = Workbook.addWorksheet("personasDB");
const hojavehiculosDB = Workbook.addWorksheet("vehiculosDB");
const hojaarmasDB = Workbook.addWorksheet("armasDB");
const hojadrogasDB = Workbook.addWorksheet("drogasDB");
const hojaResumen = Workbook.addWorksheet("Resumen");

// style header
const hojas = [
    hojaincidencias,
    hojapersonasDB,
    hojavehiculosDB,
    hojaarmasDB,
    hojadrogasDB,
    hojaResumen,
];
hojas.forEach((element) => {
    element.getRow(1).font = {
        name: "Arial",
        size: 10.5,
        bold: true,
        color: { argb: "ff000000" },
    };
    element.getRow(1).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "ffa6a6a6" },
    };
    element.getRow(1).alignment = {
        vertical: "middle",
        horizontal: "center",
    };
});

//Se generan los encabezados
hojaincidencias.columns = [
    {
        header: "Uid",
        key: "uid",
        width: 30,
    },
    {
        header: "LugarDeHechos_calle ",
        key: "calle",
        width: 30,
    },
    {
        header: "LugarDeHechos_colonia ",
        key: "colonia",
        width: 30,
    },
    {
        header: "LugarDeHechos_municipio ",
        key: "municipio",
        width: 30,
    },
    {
        header: "LugarDeHechos_estado ",
        key: "estado",
        width: 30,
    },
    {
        header: "LugarDeHechos_postal",
        key: "postal",
        width: 30,
    },
    {
        header: "LugarDeHechos_latitud",
        key: "lat",
        width: 30,
    },
    {
        header: "LugarDeHechos_Longitud",
        key: "lng",
        width: 30,
    },
    {
        header: "EntidadPersonas",
        key: "entidadPersonas",
        width: 30,
    },
    {
        header: "Capturo_nombre ",
        key: "capturo",
        width: 30,
    },
    {
        header: "Dependencia_nombre ",
        key: "dependencia",
        width: 30,
    },
    {
        header: "AreaReporte_nombre ",
        key: "areaReporte",
        width: 30,
    },
    {
        header: "TipoEvento ",
        key: "tipoEvento",
        width: 30,
    },
    {
        header: "Delito ",
        key: "delito",
        width: 30,
    },
    {
        header: "TipoDelito ",
        key: "tipoDelito",
        width: 30,
    },
    {
        header: "SubDelito ",
        key: "subDelito",
        width: 30,
    },
    {
        header: "FolioC5 ",
        key: "folioC5",
        width: 30,
    },
    {
        header: "FolioIPH ",
        key: "folioIPH",
        width: 30,
    },
    {
        header: "IphFile ",
        key: "iphFile",
        width: 30,
    },
    {
        header: "folio ",
        key: "folio",
        width: 30,
    },
    {
        header: "FechaEvento ",
        key: "fechaEvento",
        width: 30,
    },
    {
        header: "HoraEvento ",
        key: "horaEvento",
        width: 30,
    },
    {
        header: "FechaConocimiento ",
        key: "fechaConocimiento",
        width: 30,
    },
    {
        header: "HoraConocimiento ",
        key: "horaConocimiento",
        width: 30,
    },
    {
        header: "FechaRespondiente ",
        key: "fechaRespondiente",
        width: 30,
    },
    {
        header: "HoraRespondiente ",
        key: "horaRespondiente",
        width: 30,
    },
    {
        header: "PrimerRespondiente_institucion",
        key: "institucion",
        width: 30,
    },
    {
        header: "PrimerRespondiente_activo",
        key: "activo",
        width: 30,
    },
    {
        header: "PrimerRespondiente_empleado",
        key: "empleado",
        width: 30,
    },
    {
        header: "PrimerRespondiente_grado",
        key: "grado",
        width: 30,
    },
    {
        header: "PrimerRespondiente_nombre",
        key: "fechaHoraRespondiente",
        width: 30,
    },
    {
        header: "ElementosEnSitio",
        key: "elementosEnSitio",
        width: 30,
    },
    {
        header: "Narrativa",
        key: "narrativa",
        width: 30,
    },
    {
        header: "Telefonia_numero",
        key: "numero",
        width: 30,
    },
    {
        header: "Telefonia_imei",
        key: "imei",
        width: 30,
    },
    {
        header: "Telefonia_calidad",
        key: "calidad",
        width: 30,
    },
    {
        header: "Telefonia_observaciones",
        key: "observaciones",
        width: 30,
    },
    {
        header: "Objetos",
        key: "objetos",
        width: 30,
    },
    {
        header: "Dinero_cantidad",
        key: "dinerocantidad",
        width: 30,
    },
    {
        header: "Dinero_tipo",
        key: "dinerotipo",
        width: 30,
    },
    {
        header: "Dinero_calidad",
        key: "dinerocalidad",
        width: 30,
    },
    {
        header: "Dinero_observaciones",
        key: "dineroobservaciones",
        width: 30,
    },
    {
        header: "cuentas",
        key: "cuentas",
        width: 30,
    },
    {
        header: "Grupo",
        key: "grupo",
        width: 30,
    },
    {
        header: "Crp",
        key: "crp",
        width: 30,
    },
    {
        header: "Bodycam",
        key: "bodycam",
        width: 30,
    },
    {
        header: "Caso",
        key: "caso",
        width: 30,
    },
    {
        header: "FechaCaptura",
        key: "fechaCaptura",
        width: 38,
    },
    {
        header: "HoraCaptura",
        key: "horaCaptura",
        width: 38,
    },
    {
        header: "FechaActualizacion",
        key: "fechaActualizacion",
        width: 38,
    },
    {
        header: "HoraActualizacion",
        key: "horaActualizacion",
        width: 38,
    },
];

hojapersonasDB.columns = [
    {
        header: "Uid",
        key: "uid",
        width: 30,
    },
    {
        header: "IncidenteId",
        key: "incidenteId",
        width: 30,
    },
    {
        header: "Alias",
        key: "alias",
        width: 30,
    },
    {
        header: "Activo",
        key: "activo",
        width: 30,
    },
    {
        header: "FechaCreacion",
        key: "fechaCreacion",
        width: 30,
    },
    {
        header: "Nombre",
        key: "nombre",
        width: 30,
    },
    {
        header: "ApellidoP",
        key: "apellidoP",
        width: 30,
    },
    {
        header: "ApellidoM",
        key: "apellidoM",
        width: 30,
    },
    {
        header: "Edad",
        key: "edad",
        width: 30,
    },
    {
        header: "FechaNacimiento",
        key: "fechaNacimiento",
        width: 30,
    },
    {
        header: "Telefono",
        key: "telefono",
        width: 30,
    },
    {
        header: "TelefonoContacto",
        key: "telefonoContacto",
        width: 30,
    },
    {
        header: "Genero",
        key: "genero",
        width: 30,
    },
    {
        header: "EstadoCivil",
        key: "estadoCivil",
        width: 30,
    },
    {
        header: "Nacionalidad",
        key: "nacionalidad",
        width: 30,
    },
    {
        header: "Ocupacion",
        key: "ocupacion",
        width: 30,
    },
    {
        header: "Escolaridad",
        key: "escolaridad",
        width: 30,
    },
    {
        header: "NombrePadre",
        key: "nombrePadre",
        width: 30,
    },
    {
        header: "NombreMadre",
        key: "nombreMadre",
        width: 30,
    },
    {
        header: "Estatura",
        key: "estatura",
        width: 30,
    },
    {
        header: "Vestimenta",
        key: "vestimenta",
        width: 60,
    },
    {
        header: "SueldoSemanal",
        key: "sueldoSemanal",
        width: 30,
    },
    {
        header: "Domicilio_calle",
        key: "domiciliocalle",
        width: 30,
    },
    {
        header: "Domicilio_numero",
        key: "domicilionumero",
        width: 30,
    },
    {
        header: "Domicilio_colonia",
        key: "domiciliocolonia",
        width: 30,
    },
    {
        header: "Municipio",
        key: "municipio",
        width: 30,
    },
    {
        header: "Estado",
        key: "estado",
        width: 30,
    },
    {
        header: "Postal",
        key: "postal",
        width: 30,
    },
    {
        header: "Antecedentes",
        key: "antecedentes",
        width: 30,
    },
    {
        header: "Proceso",
        key: "proceso",
        width: 30,
    },
    {
        header: "Calidad_activo",
        key: "calidadactivo",
        width: 30,
    },
    {
        header: "Calidad_nombre",
        key: "calidadnombre",
        width: 30,
    },
    {
        header: "Observaciones",
        key: "observaciones",
        width: 50,
    },
    {
        header: "Identificacion",
        key: "identificacion",
        width: 30,
    },
    {
        header: "SeñasParticulares",
        key: "señasParticulares",
        width: 30,
    },
];
hojapersonasDB.getColumn(9).value = "12345";

hojavehiculosDB.columns = [
    {
        header: "Uid",
        key: "uid",
        width: 30,
    },
    {
        header: "IncidenteId",
        key: "incidenteId",
        width: 30,
    },
    {
        header: "Activo",
        key: "activo",
        width: 30,
    },
    {
        header: "Calidad_nombre",
        key: "calidanombre",
        width: 30,
    },
    {
        header: "Tipo_descripcion",
        key: "tipodescripcion",
        width: 30,
    },
    {
        header: "Modelo",
        key: "modelo",
        width: 30,
    },
    {
        header: "Marca_descripcion",
        key: "marcadescripcion",
        width: 30,
    },
    {
        header: "Submarca_descripcion",
        key: "submarcadescripcion",
        width: 30,
    },
    {
        header: "Color",
        key: "color",
        width: 30,
    },
    {
        header: "Placa",
        key: "placa",
        width: 30,
    },
    {
        header: "Serie",
        key: "serie",
        width: 30,
    },
    {
        header: "Motor",
        key: "motor",
        width: 30,
    },
    {
        header: "Niv",
        key: "niv",
        width: 30,
    },
    {
        header: "Propietario",
        key: "propietario",
        width: 30,
    },
    {
        header: "Observaciones",
        key: "observaciones",
        width: 30,
    },
    {
        header: "FechaCreacion",
        key: "fechaCreacion",
        width: 30,
    },
];
hojaarmasDB.columns = [
    {
        header: "Uid",
        key: "uid",
        width: 30,
    },
    {
        header: "incidenteId",
        key: "incidenteId",
        width: 30,
    },
    {
        header: "Activo",
        key: "activo",
        width: 30,
    },
    {
        header: "Cantidad",
        key: "cantidad",
        width: 30,
    },
    {
        header: "Tipo_descripcion",
        key: "tipo",
        width: 30,
    },
    {
        header: "Calibre_descripcion",
        key: "calibre",
        width: 30,
    },
    {
        header: "matricula",
        key: "matricula",
        width: 30,
    },
    {
        header: "fabricante",
        key: "fabricante",
        width: 30,
    },
    {
        header: "noSerie",
        key: "noSerie",
        width: 30,
    },
    {
        header: "modelo",
        key: "modelo",
        width: 30,
    },
    {
        header: "Calidad_nombre",
        key: "calidad",
        width: 30,
    },
    {
        header: "fechaCreacion",
        key: "fechaCreacion",
        width: 30,
    },

    {
        header: "observaciones",
        key: "observaciones",
        width: 30,
    },
];
hojadrogasDB.columns = [
    {
        header: "Uid",
        key: "uid",
        width: 30,
    },
    {
        header: "IncidenteId",
        key: "incidenteId",
        width: 30,
    },
    {
        header: "activo",
        key: "activo",
        width: 30,
    },
    {
        header: "Tipo",
        key: "tipo",
        width: 30,
    },
    {
        header: "Cantidad",
        key: "cantidad",
        width: 30,
    },
    {
        header: "Unidad",
        key: "unidad",
        width: 30,
    },
    {
        header: "Calidad",
        key: "calidad",
        width: 30,
    },
    {
        header: "Observaciones",
        key: "observaciones",
        width: 30,
    },
    {
        header: "fechaCreacion",
        key: "fechaCreacion",
        width: 30,
    },
];

hojaResumen.columns = [
    {
        header: "Fechas",
        key: "fechas",
        width: 30,
    },
    {
        header: "Modulo",
        key: "module",
        width: 30,
    },
    {
        header: "Total",
        key: "total",
        width: 10,
    },
];
module.exports = {
    Workbook,
    hojaincidencias,
    hojapersonasDB,
    hojavehiculosDB,
    hojaarmasDB,
    hojadrogasDB,
    hojaResumen,
};
