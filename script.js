let invitados = [];


// ======================================
// CARGAR EXCEL
// ======================================

window.onload = async function () {

    try {

        const respuesta = await fetch(
            'Invitados.xlsx?v=' + new Date().getTime()
        );

        const data = await respuesta.arrayBuffer();

        const workbook = XLSX.read(data, {
            type: 'array'
        });

        const hoja =
            workbook.Sheets[workbook.SheetNames[0]];

        invitados = XLSX.utils.sheet_to_json(hoja);

        // Agregar estado
        invitados = invitados.map(persona => {

            return {

                ...persona,

                estado: false,

                nombreCompleto: `
                    ${persona["Nombre"] || ""}
                    ${persona["Apellido P"] || ""}
                    ${persona["Apellido M"] || ""}
                `
                .replace(/\s+/g, ' ')
                .trim()

            };

        });

        mostrarInvitados(invitados);

        actualizarContadores();

    } catch (error) {

        console.error(error);

        alert("Error cargando Excel");

    }

};


// ======================================
// MOSTRAR INVITADOS
// ======================================

function mostrarInvitados(lista) {

    const tabla =
        document.getElementById("tablaInvitados");

    tabla.innerHTML = "";

    lista.forEach((persona, index) => {

        const fila =
            document.createElement("tr");

        fila.innerHTML = `

            <td>
                ${persona["Cargo/Grado"] || ""}
            </td>

            <td>
                ${persona.nombreCompleto || ""}
            </td>

            <td>
                ${persona["Sector"] || ""}
            </td>

            <td>
                ${persona["Asiento"] || ""}
            </td>

            <td>

                ${
                    persona.estado
                    ?
                    `<button
                        class="btn btn-danger"
                        onclick="deshacerIngreso(${index})"
                    >
                        Deshacer
                    </button>`
                    :
                    `<button
                        class="btn btn-success"
                        onclick="marcarIngreso(${index})"
                    >
                        Ingresó
                    </button>`
                }

            </td>

        `;

        tabla.appendChild(fila);

    });

}


// ======================================
// MARCAR INGRESO
// ======================================

function marcarIngreso(index) {

    invitados[index].estado = true;

    mostrarInvitados(invitados);

    actualizarContadores();

}


// ======================================
// DESHACER INGRESO
// ======================================

function deshacerIngreso(index) {

    invitados[index].estado = false;

    mostrarInvitados(invitados);

    actualizarContadores();

}


// ======================================
// CONTADORES
// ======================================

function actualizarContadores() {

    const total = invitados.length;

    const ingresados = invitados.filter(
        persona => persona.estado
    ).length;

    const pendientes = total - ingresados;

    document.getElementById(
        "totalInvitados"
    ).textContent = total;

    document.getElementById(
        "ingresados"
    ).textContent = ingresados;

    document.getElementById(
        "pendientes"
    ).textContent = pendientes;

}


// ======================================
// BUSCADOR
// ======================================

document.addEventListener(
    "DOMContentLoaded",
    () => {

        const buscador =
            document.getElementById("buscador");

        buscador.addEventListener(
            "keyup",
            () => {

                const texto =
                    buscador.value.toLowerCase();

                const filtrados =
                    invitados.filter(persona => {

                        const contenido = `
                            ${persona["Cargo/Grado"] || ""}
                            ${persona.nombreCompleto || ""}
                            ${persona["Sector"] || ""}
                            ${persona["Asiento"] || ""}
                        `
                        .toLowerCase();

                        return contenido.includes(texto);

                    });

                mostrarInvitados(filtrados);

            }
        );

    }
);