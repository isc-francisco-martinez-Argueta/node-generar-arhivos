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
        size: 10,
        bold: true,
        color: { argb: "ff595959" },
    };
    element.getRow(1).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "ffd9d9d9" },
    };
    element.getRow(1).alignment = {
        vertical: "middle",
        horizontal: "center",
    };
});

//Se generan los encabezados
hojaincidencias.columns = [
    {
        header: "uid",
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
        header: "LugarDeHechos_postal ",
        key: "postal",
        width: 30,
    },
    {
        header: "LugarDeHechos_latitud ",
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
        header: "FechaHoraEvento ",
        key: "fechaHoraEvento",
        width: 30,
    },
    {
        header: "FechaHoraConocimiento ",
        key: "fechaHoraConocimiento",
        width: 30,
    },
    {
        header: "FechaHoraRespondiente ",
        key: "fechaHoraRespondiente",
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
        header: "FechaHoraCaptura",
        key: "fechaHoraCaptura",
        width: 30,
    },
    {
        header: "FechaHoraActualizacion",
        key: "fechaHoraActualizacion",
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
