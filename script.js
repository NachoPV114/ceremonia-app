let invitados = [];
let ingresados = JSON.parse(localStorage.getItem("ingresados")) || [];

// Cargar Excel automáticamente
window.onload = async function () {

    const respuesta = await fetch('Invitados.xlsx');
    const data = await respuesta.arrayBuffer();

    const workbook = XLSX.read(data, {
        type: 'array'
    });

    const hoja = workbook.Sheets[workbook.SheetNames[0]];

    invitados = XLSX.utils.sheet_to_json(hoja);

    invitados = invitados.map(persona => {

        const yaIngreso = ingresados.includes(
            persona["Nombre Completo"]
        );

        return {
            ...persona,
            estado: yaIngreso
        };

    });

    mostrarInvitados(invitados);
    actualizarContadores();

};

// Mostrar invitados
function mostrarInvitados(lista) {

    const tabla = document.getElementById("tablaInvitados");

    tabla.innerHTML = "";

    lista.forEach((persona, index) => {

        const fila = document.createElement("tr");

        fila.innerHTML = `

            <td>${persona["Grado"] || ""}</td>

            <td>${persona["Nombre Completo"] || ""}</td>

            <td>${persona["Sector"] || ""}</td>

            <td>${persona["Asiento"] || ""}</td>

            <td>

                ${
                    persona.estado
                    ?
                    `<button class="btn btn-danger"
                        onclick="deshacerIngreso(${index})">
                        Deshacer
                    </button>`
                    :
                    `<button class="btn btn-success"
                        onclick="marcarIngreso(${index})">
                        Marcar ingreso
                    </button>`
                }

            </td>

        `;

        tabla.appendChild(fila);

    });

}

// Marcar ingreso
function marcarIngreso(index) {

    invitados[index].estado = true;

    const nombre = invitados[index]["Nombre Completo"];

    if (!ingresados.includes(nombre)) {
        ingresados.push(nombre);
    }

    localStorage.setItem(
        "ingresados",
        JSON.stringify(ingresados)
    );

    mostrarInvitados(invitados);
    actualizarContadores();

}

// Deshacer ingreso
function deshacerIngreso(index) {

    invitados[index].estado = false;

    const nombre = invitados[index]["Nombre Completo"];

    ingresados = ingresados.filter(
        item => item !== nombre
    );

    localStorage.setItem(
        "ingresados",
        JSON.stringify(ingresados)
    );

    mostrarInvitados(invitados);
    actualizarContadores();

}

// Actualizar contadores
function actualizarContadores() {

    const total = invitados.length;

    const totalIngresados = invitados.filter(
        persona => persona.estado
    ).length;

    const pendientes = total - totalIngresados;

    document.getElementById("totalInvitados").textContent = total;

    document.getElementById("ingresados").textContent = totalIngresados;

    document.getElementById("pendientes").textContent = pendientes;

}

// Buscador
document.getElementById("buscador").addEventListener("input", function () {

    const texto = this.value.toLowerCase();

    const filtrados = invitados.filter(persona => {

        return Object.values(persona)
            .join(" ")
            .toLowerCase()
            .includes(texto);

    });

    mostrarInvitados(filtrados);

});