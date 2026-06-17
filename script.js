let invitados = [];
let invitadosFiltrados = [];
let ingresos = {};

window.onload = async function () {

    try {

        const respuesta =
            await fetch("Invitados.xlsx");

        const data =
            await respuesta.arrayBuffer();

        const workbook =
            XLSX.read(data, {
                type: "array"
            });

        const nombreHoja =
            workbook.SheetNames[0];

        const hoja =
            workbook.Sheets[nombreHoja];

        invitados =
            XLSX.utils.sheet_to_json(hoja);

        invitadosFiltrados =
            [...invitados];

        await cargarIngresosFirebase();

        actualizarContadores();

        mostrarInvitados(
            invitadosFiltrados
        );

    }

    catch (error) {

        console.error(error);

        alert(
            "Error cargando Excel"
        );

    }

};

function obtenerNombre(persona) {

    const nombreCompleto =
        String(
            persona["NOMBRE COMPLETO"] || ""
        ).trim();

    if (nombreCompleto !== "") {
        return nombreCompleto;
    }

    const nombre =
        String(
            persona["NOMBRE"] || ""
        ).trim();

    const apellidoP =
        String(
            persona["APE. P."] || ""
        ).trim();

    const apellidoM =
        String(
            persona["APE. M."] || ""
        ).trim();

    const nombreArmado =
        `${nombre} ${apellidoP} ${apellidoM}`
        .replace(/\s+/g, " ")
        .trim();

    if (nombreArmado !== "") {
        return nombreArmado;
    }

    return "Sin nombre";

}

function actualizarContadores() {

    const ingresados =
        Object.keys(
            ingresos
        ).length;

    document.getElementById(
        "totalInvitados"
    ).textContent =
        invitados.length;

    document.getElementById(
        "totalIngresados"
    ).textContent =
        ingresados;

    document.getElementById(
        "totalPendientes"
    ).textContent =
        invitados.length - ingresados;

}

document.addEventListener(
    "DOMContentLoaded",
    () => {

        const buscador =
            document.getElementById(
                "buscador"
            );

        buscador.addEventListener(
            "input",
            () => {

                const texto =
                    buscador.value
                    .toLowerCase();

                invitadosFiltrados =
                    invitados.filter(
                        persona => {

                            const contenido = `
                                ${persona["NOMBRE COMPLETO"] || ""}
                                ${persona["NOMBRE"] || ""}
                                ${persona["APE. P."] || ""}
                                ${persona["APE. M."] || ""}
                                ${persona["CARGO"] || ""}
                                ${persona["TRIBUNA"] || ""}
                                ${persona["ASIENTO"] || ""}
                            `
                                .toLowerCase();

                            return contenido.includes(
                                texto
                            );

                        }
                    );

                mostrarInvitados(
                    invitadosFiltrados
                );

            }
        );

    }
);

async function guardarEnFirebase(id) {

    try {

        const ref =
            window.firebaseFirestore.doc(
                window.db,
                "ingresos",
                String(id)
            );

        await window.firebaseFirestore.setDoc(
            ref,
            {
                ingresado:
                    ingresos[id] === true,

                fecha:
                    new Date()
                        .toISOString()
            }
        );

    }

    catch (error) {

        console.error(
            error
        );

    }

}

async function cargarIngresosFirebase() {

    window.firebaseFirestore.onSnapshot(

        window.coleccionIngresos,

        (snapshot) => {

            ingresos = {};

            snapshot.forEach((docu) => {

                const datos = docu.data();

                if (datos.ingresado === true) {

                    ingresos[
                        docu.id
                    ] = true;

                }

            });

            actualizarContadores();

            mostrarInvitados(
                invitadosFiltrados
            );

        }

    );

}