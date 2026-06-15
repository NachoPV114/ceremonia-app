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
            workbook.SheetNames.find(
                h => h.toLowerCase() === "maestro"
            );

        const hoja =
            workbook.Sheets[nombreHoja];

        invitados =
            XLSX.utils.sheet_to_json(hoja);

        console.log(invitados[0]);

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


// ============================
// NOMBRE COMPLETO
// ============================

function obtenerNombre(persona) {

    return (
        persona["NOMBRE COMPLETO"] ||
        (
            `${persona["NOMBRE"] || ""} ${
                persona["APE. P."] || ""
            } ${
                persona["APE. M."] || ""
            }`
        ).trim()
    );

}


// ============================
// UBICACION
// ============================

function obtenerUbicacion(persona) {

    const tribuna =
        persona["TRIBUNA"] || "";

    const asiento =
        persona["ASIENTO"] || "";

    return `${tribuna} - ${asiento}`;

}


// ============================
// MOSTRAR TABLA
// ============================

function mostrarInvitados(lista) {

    const tabla =
        document.getElementById(
            "tablaInvitados"
        );

    tabla.innerHTML = "";

    lista.forEach(persona => {

        const id =
            persona["N°"];

        const ingresado =
            ingresos[id] === true;

        const fila =
            document.createElement("tr");

        fila.innerHTML = `

            <td>

                <button
                    class="btn ${ingresado ? 'btn-danger' : 'btn-success'} btn-sm"
                    onclick="cambiarEstado(${id})">

                    ${ingresado ? 'Deshacer' : 'Marcar ingreso'}

                </button>

            </td>

            <td>
                ${obtenerNombre(persona)}
            </td>

            <td>
                ${persona["Cargo"] || ""}
            </td>

            <td>
                ${persona["Empresa/Universidad/Otro"] || ""}
            </td>

            <td>
                ${persona["Categoría"] || ""}
            </td>

            <td>
                ${obtenerUbicacion(persona)}
            </td>

        `;

        tabla.appendChild(
            fila
        );

    });

}
function cambiarEstado(id) {

    if (ingresos[id]) {

        delete ingresos[id];

    }

    else {

        ingresos[id] = true;

    }
    guardarEnFirebase(id);

 localStorage.setItem(
        "ingresos",
        JSON.stringify(
            ingresos
        )
    );

    actualizarContadores();

    mostrarInvitados(
        invitadosFiltrados
    );

}


// ============================
// CONTADORES
// ============================
const guardados =
    localStorage.getItem(
        "ingresos"
    );

if (guardados) {

    ingresos =
        JSON.parse(
            guardados
        );

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


// ============================
// BUSCADOR
// ============================

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
                                ${persona["Apellido Paterno"] || ""}
                                ${persona["Apellido Materno"] || ""}
                                ${persona["Cargo"] || ""}
                                ${persona["Empresa/Universidad/Otro"] || ""}
                                ${persona["Categoría"] || ""}
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

        console.log(
            "Guardado Firebase:",
            id
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
                    ingresos[docu.id] = true;
                }

            });

            console.log(
                "Ingresos sincronizados:",
                Object.keys(ingresos).length
            );

            actualizarContadores();

            mostrarInvitados(
                invitadosFiltrados
            );

        }

    );

}
