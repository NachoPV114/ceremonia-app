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
                h => h.toLowerCase() === "consolidado"
            );

        const hoja =
            workbook.Sheets[nombreHoja];

        invitados =
            XLSX.utils.sheet_to_json(hoja);

        invitadosFiltrados =
            [...invitados];

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

    return `
        ${persona["Prefijo"] || ""}
        ${persona["Nombre"] || ""}
        ${persona["Apellido Paterno"] || ""}
        ${persona["Apellido Materno"] || ""}
    `
        .replace(/\s+/g, " ")
        .trim();

}


// ============================
// UBICACION
// ============================

function obtenerUbicacion(persona) {

    const tribuna =
        persona["Tribuna"] || "";

    const asiento =
        persona["Asiento"] || "";

    return `
        ${tribuna}
        Asiento ${asiento}
    `
        .replace(/\s+/g, " ")
        .trim();

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
            persona["Nro."];

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