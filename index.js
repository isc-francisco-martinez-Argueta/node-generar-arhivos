const exceljs = require("exceljs");
const modeloxls = require("./modeloXLS.json");
const fs = require("fs");
const createExcel = async () => {
    // Crear un archivo excel
    const Workbook = new exceljs.Workbook();

    const hojaincidencias = Workbook.addWorksheet("incidencias");
    const hojapersonasDB = Workbook.addWorksheet("personasDB");
    const hojavehiculosDB = Workbook.addWorksheet("vehiculosDB");
    const hojaarmasDB = Workbook.addWorksheet("armasDB");
    const hojadrogasDB = Workbook.addWorksheet("drogasDB");

    hojaincidencias.columns = [
        {
            header: "uid",
            key: "uid",
            width: 30,
        },
        {
            header: "lugarDeHechos calle ",
            key: "calle",
            width: 30,
        },
        {
            header: "lugarDeHechos colonia ",
            key: "colonia",
            width: 30,
        },
        {
            header: "lugarDeHechos municipio ",
            key: "municipio",
            width: 30,
        },
        {
            header: "lugarDeHechos estado ",
            key: "estado",
            width: 30,
        },
        {
            header: "lugarDeHechos postal ",
            key: "postal",
            width: 30,
        },
        {
            header: "lugarDeHechos lat ",
            key: "lat",
            width: 30,
        },
        {
            header: "lugarDeHechos lng ",
            key: "lng",
            width: 30,
        },
        {
            header: "capturo nombre ",
            key: "capturo",
            width: 30,
        },
        {
            header: "dependencia nombre ",
            key: "dependencia",
            width: 30,
        },
        {
            header: "areaReporte nombre ",
            key: "areaReporte",
            width: 30,
        },
        {
            header: "tipoEvento ",
            key: "tipoEvento",
            width: 30,
        },
        {
            header: "delito ",
            key: "delito",
            width: 30,
        },
        {
            header: "tipoDelito ",
            key: "tipoDelito",
            width: 30,
        },
        {
            header: "subDelito ",
            key: "subDelito",
            width: 30,
        },
        {
            header: "folioC5 ",
            key: "folioC5",
            width: 30,
        },
        {
            header: "folioIPH ",
            key: "folioIPH",
            width: 30,
        },
        {
            header: "iphFile ",
            key: "iphFile",
            width: 30,
        },
        {
            header: "folio ",
            key: "folio",
            width: 30,
        },
    ];
    let object = JSON.parse(fs.readFileSync("modeloXLS.json", "utf8"));

    try {
        await object.incidencias.map((value, idx) => {
            hojaincidencias.addRow({
                uid: value.uid,
                calle: value.lugarDeHechos.calle,
                colonia: value.lugarDeHechos.colonia,
                municipio: value.lugarDeHechos.municipio,
                estado: value.lugarDeHechos.estado,
                postal: value.lugarDeHechos.postal,
                lat: value.lugarDeHechos.lat,
                lng: value.lugarDeHechos.lng,
                capturo: value.capturo.nombre,
                dependencia: value.dependencia.nombre,
                areaReporte: value.areaReporte.nombre,
                tipoEvento: value.tipoEvento,
                delito: value.delito,
                tipoDelito: value.tipoDelito,
                subDelito: value.subDelito,
                folioC5: value.folioC5,
                folioIPH: value.folioIPH,
                iphFile: value.iphFile,
                folio: value.folio,
            });
        });
        await Workbook.xlsx.writeFile("./salida/modeloXLS.xlsx");
        return `modeloXLS.xlsx`;
    } catch (err) {
        throw err;
    }
};

createExcel()
    .then((nombreArchivo) => console.log(nombreArchivo, "Creado"))
    .catch((err) => console.log(err));
