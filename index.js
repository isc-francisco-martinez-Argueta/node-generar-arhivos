const modeloxls = require("./modeloXLS.json");
const fs = require("fs");
const moment = require("moment");
moment.locale("es-mx");
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
        await model.incidentesDB.map((incidencia, idx) => {
            incidencia.telefonia.forEach((incidenciatelefonia) => {
                incidencia.dinero.forEach((incidenciadinero) => {
                    hojaincidencias.addRow({
                        uid: incidencia.uid,
                        calle: incidencia.lugarDeHechos.calle,
                        colonia: incidencia.lugarDeHechos.colonia,
                        municipio: incidencia.lugarDeHechos.municipio,
                        estado: incidencia.lugarDeHechos.estado,
                        postal: incidencia.lugarDeHechos.postal,
                        lat: incidencia.lugarDeHechos.lat,
                        lng: incidencia.lugarDeHechos.lng,
                        entidadPersonas: Object.values(
                            incidencia.entidadPersonas
                        ).toString(),
                        capturo: incidencia.capturo.nombre,
                        dependencia: incidencia.dependencia.nombre,
                        areaReporte: incidencia.areaReporte.nombre,
                        tipoEvento: incidencia.tipoEvento,
                        delito: incidencia.delito,
                        tipoDelito: incidencia.tipoDelito,
                        subDelito: incidencia.subDelito,
                        folioC5: incidencia.folioC5,
                        folioIPH: incidencia.folioIPH,
                        iphFile: incidencia.iphFile,
                        folio: incidencia.folio,
                        fechaEvento: incidencia.fechaHoraEvento
                            ? moment(incidencia.fechaHoraEvento).format("L")
                            : "",
                        horaEvento: moment(incidencia.fechaHoraEvento).format(
                            "hh:mm a"
                        ),
                        fechaConocimiento: incidencia.fechaHoraConocimiento
                            ? moment(incidencia.fechaHoraConocimiento).format(
                                  "L"
                              )
                            : "",
                        horaConocimiento: incidencia.fechaHoraConocimiento
                            ? moment(incidencia.fechaHoraConocimiento).format(
                                  "hh:mm a"
                              )
                            : "",
                        fechaRespondiente: incidencia.fechaHoraRespondiente
                            ? moment(incidencia.fechaHoraRespondiente).format(
                                  "L"
                              )
                            : "",
                        horaRespondiente: incidencia.fechaHoraRespondiente
                            ? moment(incidencia.fechaHoraRespondiente).format(
                                  "hh:mm a"
                              )
                            : "",
                        institucion: incidencia.primerRespondiente.institucion,
                        activo: incidencia.primerRespondiente.activo
                            ? "Activo"
                            : "Inactivo",
                        empleado: incidencia.primerRespondiente.empleado,
                        grado: incidencia.primerRespondiente.grado,
                        nombre: incidencia.primerRespondiente.nombre,
                        elementosEnSitio:
                            incidencia.elementosEnSitio.toString(),
                        narrativa: incidencia.narrativa,
                        numero: incidenciatelefonia.datosAdicionales.telefonia
                            .numero,
                        imei: incidenciatelefonia.datosAdicionales
                            .telefoniaimei,
                        calidad: incidenciatelefonia.calidad,
                        observaciones: incidenciatelefonia.observaciones,
                        objetos: incidencia.objetos.toString(),
                        dinerocantidad:
                            incidenciadinero.datosAdicionales.dinero.cantidad,
                        dinerotipo:
                            incidenciadinero.datosAdicionales.dinero.tipo,
                        dinerocalidad: incidenciadinero.calidad,
                        dineroobservaciones: incidenciadinero.observaciones,
                        cuentas: incidencia.cuentas.toString(),
                        grupo: incidencia.grupo,
                        crp: incidencia.crp,
                        bodycam: incidencia.bodycam,
                        caso: incidencia.caso,
                        fechaCaptura: incidencia.fechaHoraCaptura
                            ? moment(incidencia.fechaHoraCaptura).format("L ")
                            : "",
                        horaCaptura: incidencia.fechaHoraCaptura
                            ? moment(incidencia.fechaHoraCaptura).format(
                                  "hh:mm a"
                              )
                            : "",
                        fechaActualizacion: incidencia.fechaHoraActualizacion
                            ? moment(incidencia.fechaHoraActualizacion).format(
                                  "L"
                              )
                            : "",
                        horaActualizacion: incidencia.fechaHoraActualizacion
                            ? moment(incidencia.fechaHoraActualizacion).format(
                                  "hh:mm a"
                              )
                            : "",
                    });
                });
            });
            countincidencia = countincidencia + 1;
        });

        await model.personasDB.map((persona, idx) => {
            hojapersonasDB.addRow({
                uid: persona.uid,
                incidenteId: persona.incidenteId,
                alias: Object.values(persona.alias).toString(),
                activo: persona.activo ? "Si" : "No",
                fechaCreacion: moment(persona.fechaCreacion).format("L"),
                nombre: persona.nombre,
                apellidoP: persona.apellidoP,
                apellidoM: persona.apellidoM,
                edad: persona.edad,
                fechaNacimiento: moment(persona.fechaNacimiento).format("L"),
                telefono: persona.telefono,
                telefonoContacto: persona.telefonoContacto,
                genero: persona.genero.descripcion,
                estadoCivil: persona.estadoCivil.descripcion,
                nacionalidad: persona.nacionalidad.descripcion,
                ocupacion: persona.ocupacion.descripcion,
                escolaridad: persona.escolaridad.descripcion,
                nombrePadre: persona.nombrePadre,
                nombreMadre: persona.nombreMadre,
                estatura: persona.estatura,
                vestimenta: persona.vestimenta,
                sueldoSemanal: persona.sueldoSemanal,
                domiciliocalle: persona.domicilio.calle,
                domicilionumero: persona.domicilio.numero,
                domiciliocolonia: persona.domicilio.colonia,
                municipio: persona.domicilio.municipio.descripcion,
                estado: persona.domicilio.estado.descripcion,
                postal: persona.domicilio.postal,
                antecedentes: Object.values(persona.antecedentes).toString(),
                proceso: persona.proceso,
                calidadactivo: persona.calidad.activo ? "Activo" : "Inactivo",
                calidadnombre: persona.calidad.nombre,
                observaciones: persona.observaciones,
                identificacion: Object.values(
                    persona.identificacion
                ).toString(),
                señasParticulares: Object.values(
                    persona.señasParticulares
                ).toString(),
            });
            countpersonasDB = countpersonasDB + 1;
        });

        await model.vehiculosDB.map((vehiculo, idx) => {
            hojavehiculosDB.addRow({
                uid: vehiculo.uid,
                incidenteId: vehiculo.incidenteId,
                activo: vehiculo.activo ? "Activo" : "Inactivo",
                calidanombre: vehiculo.calidad.nombre,
                tipodescripcion: vehiculo.tipo.descripcion,
                modelo: vehiculo.modelo,
                marcadescripcion: vehiculo.marca.descripcion,
                submarcadescripcion: vehiculo.submarca.descripcion,
                color: vehiculo.color,
                placa: vehiculo.placa,
                serie: vehiculo.serie,
                motor: vehiculo.motor,
                niv: vehiculo.niv,
                propietario: vehiculo.propietario,
                observaciones: vehiculo.observaciones,
                fechaCreacion: moment(vehiculo.fechaCreacion).format("L"),
            });
            countvehiculosDB = countvehiculosDB + 1;
        });
        await model.armasDB.map((arma, idx) => {
            hojaarmasDB.addRow({
                uid: arma.uid,
                incidenteId: arma.incidenteId,
                activo: arma.activo ? "Activo" : "Inactivo",
                cantidad: arma.cantidad,
                tipo: arma.tipo.descripcion,
                calibre: arma.calibre.descripcion,
                matricula: arma.matricula,
                fabricante: arma.fabricante,
                noSerie: arma.noSerie,
                modelo: arma.modelo,
                calidad: arma.calidad.nombre,
                fechaCreacion: moment(arma.fechaCreacion).format("L"),
                observaciones: arma.observaciones,
            });
            countarmasDB = countarmasDB + 1;
        });
        await model.drogasDB.map((droga, idx) => {
            hojadrogasDB.addRow({
                uid: droga.uid,
                incidenteId: droga.incidenteId,
                activo: droga.activo ? "Activo" : "inactivo",
                tipo: droga.tipo.descripcion,
                cantidad: droga.cantidad,
                unidad: droga.unidad,
                calidad: droga.calidad.nombre,
                observaciones: droga.observaciones,
                fechaCreacion: moment(droga.fechaCreacion).format("L"),
            });
            countdrogasDB = countdrogasDB + 1;
        });
        hojaResumen.addRow({
            fechas: "fechas",
            module: "Incidencias",
            total: countincidencia,
        });
        hojaResumen.addRow({
            fechas: "fechas",
            module: "Personas",
            total: countpersonasDB,
        });
        hojaResumen.addRow({
            fechas: "fechas",
            module: "Vehiculos",
            total: countvehiculosDB,
        });
        hojaResumen.addRow({
            fechas: "fechas",
            module: "Armas",
            total: countarmasDB,
        });
        hojaResumen.addRow({
            fechas: "fechas",
            module: "Drogas",
            total: countdrogasDB,
        });

        await Workbook.xlsx.writeFile("./salida/modeloXLS.xlsx");
        return `modeloXLS.xlsx`;
    } catch (err) {
        throw err;
    }
};

createExcel(modeloxls)
    .then((nombreArchivo) => console.log(nombreArchivo, "Creado"))
    .catch((err) => console.log(err));
