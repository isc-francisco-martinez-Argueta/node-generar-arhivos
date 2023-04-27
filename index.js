const modeloxls = require("./modeloXLS.json");
const fs = require("fs");
const {
    Workbook,
    hojaincidencias,
    hojapersonasDB,
    hojavehiculosDB,
    hojaarmasDB,
    hojadrogasDB,
    hojaResumen,
} = require("./helpers/hojas");

const createExcel = async (model) => {
    let countincidencia = 0,
        countpersonasDB = 0,
        countvehiculosDB = 0,
        countarmasDB = 0,
        countdrogasDB = 0;

    try {
        await model.incidencias.map((value, idx) => {
            value.telefonia.forEach((element) => {
                value.dinero.forEach((element2) => {
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
                        fechaHoraEvento: value.fechaHoraEvento,
                        fechaHoraConocimiento: value.fechaHoraConocimiento,
                        fechaHoraConocimiento: value.fechaHoraConocimiento,
                        institucion: value.primerRespondiente.institucion,
                        activo: value.primerRespondiente.activo
                            ? "Activo"
                            : "Inactivo",
                        empleado: value.primerRespondiente.empleado,
                        grado: value.primerRespondiente.grado,
                        nombre: value.primerRespondiente.nombre,
                        elementosEnSitio: value.elementosEnSitio,
                        narrativa: value.narrativa,
                        numero: element.datosAdicionales.telefonia.numero,
                        imei: element.datosAdicionales.telefoniaimei,
                        calidad: element.calidad,
                        observaciones: element.observaciones,
                        objetos: value.objetos,
                        dinerocantidad: `$ ${element2.datosAdicionales.dinero.cantidad}`,
                        dinerotipo: element2.datosAdicionales.dinero.tipo,
                        dinerocalidad: element2.calidad,
                        dineroobservaciones: element2.observaciones,
                        cuentas: value.cuentas,
                        grupo: value.grupo,
                        crp: value.crp,
                        bodycam: value.bodycam,
                        caso: value.caso,
                        fechaHoraCaptura: value.fechaHoraCaptura,
                        fechaHoraActualizacion: value.fechaHoraActualizacion,
                    });
                    console.log(element2.observaciones);
                });
            });

            countincidencia = countincidencia + 1;
        });
        hojaResumen.addRow({
            fechas: "fechas",
            module: "incidencias",
            total: countincidencia,
        });

        console.log(countincidencia);
        await Workbook.xlsx.writeFile("./salida/modeloXLS.xlsx");
        return `modeloXLS.xlsx`;
    } catch (err) {
        throw err;
    }
};

createExcel(modeloxls)
    .then((nombreArchivo) => console.log(nombreArchivo, "Creado"))
    .catch((err) => console.log(err));
