let invitados = [];

let ingresados = JSON.parse(
    localStorage.getItem("ingresados")
) || [];


// ===============================
// CARGAR EXCEL AUTOMÁTICAMENTE
// ===============================

window.onload = async function () {

    const respuesta = await fetch("Invitados.xlsx");

    const data = await respuesta.arrayBuffer();

    const workbook = XLSX.read(data, {
        type: "array"
    });

    const hoja = workbook.Sheets[workbook.SheetNames[0]];

    invitados = XLSX.utils.sheet_to_json(hoja);

    mostrarInvitados(invitados);

    actualizarContadores();
};


// ===============================
// MOSTRAR INVITADOS
// ===============================

function mostrarInvitados(lista) {

    const tabla = document.getElementById("tablaInvitados");

    tabla.innerHTML = "";

    lista.forEach((persona) => {

        const fila = document.createElement("tr");

        const idPersona =
            `${persona["Nombre"] || ""}${persona["Apellido P"] || ""}`;


        fila.innerHTML = `
            <td>${persona["Cargo/Grado"] || ""}</td>

            <td>
                ${persona["Nombre"] || ""} 
                ${persona["Apellido P"] || ""} 
                ${persona["Apellido M"] || ""}
            </td>

            <td>${persona["Sector"] || ""}</td>

            <td>${persona["Asiento"] || ""}</td>

            <td>
                <button
                    class="btn ${
                        ingresados.includes(idPersona)
                            ? "btn-danger"
                            : "btn-success"
                    }"

                    onclick="marcarIngreso('${idPersona}')"
                >
                    ${
                        ingresados.includes(idPersona)
                            ? "Deshacer"
                            : "Ingresó"
                    }
                </button>
            </td>
        `;

        tabla.appendChild(fila);

    });

}


// ===============================
// MARCAR / DESHACER INGRESO
// ===============================

function marcarIngreso(idPersona) {

    if (ingresados.includes(idPersona)) {

        ingresados = ingresados.filter(
            item => item !== idPersona
        );

    } else {

        ingresados.push(idPersona);

    }

    localStorage.setItem(
        "ingresados",
        JSON.stringify(ingresados)
    );

    mostrarInvitados(invitados);

    actualizarContadores();
}


// ===============================
// CONTADORES
// ===============================

function actualizarContadores() {

    document.getElementById("totalInvitados").textContent =
        invitados.length;

    document.getElementById("totalIngresados").textContent =
        ingresados.length;

    document.getElementById("totalPendientes").textContent =
        invitados.length - ingresados.length;
}


// ===============================
// BUSCADOR
// ===============================

document
    .getElementById("buscador")
    .addEventListener("input", function () {

        const texto = this.value.toLowerCase();

        const filtrados = invitados.filter(persona => {

            const nombreCompleto = `
                ${persona["Nombre"] || ""}
                ${persona["Apellido P"] || ""}
                ${persona["Apellido M"] || ""}
                ${persona["Cargo/Grado"] || ""}
            `
                .toLowerCase();

            return nombreCompleto.includes(texto);

        });

        mostrarInvitados(filtrados);

    });