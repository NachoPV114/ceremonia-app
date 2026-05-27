import {
    collection,
    getDocs,
    addDoc,
    deleteDoc,
    doc,
    query,
    where
}
from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

let invitados = [];

let ingresados = [];


// ===============================
// CARGAR TODO
// ===============================

window.onload = async function () {

    // Cargar Excel
    const respuesta =
        await fetch("Invitados.xlsx");

    const data =
        await respuesta.arrayBuffer();

    const workbook =
        XLSX.read(data, {
            type: "array"
        });

    const hoja =
        workbook.Sheets[
            workbook.SheetNames[0]
        ];

    invitados =
        XLSX.utils.sheet_to_json(hoja);

    // Cargar ingresos Firebase
    await cargarIngresados();

    // Mostrar
    mostrarInvitados(invitados);

    actualizarContadores();

};


// ===============================
// CARGAR INGRESADOS FIREBASE
// ===============================

async function cargarIngresados() {

    ingresados = [];

    const snapshot = await getDocs(
        collection(window.db, "ingresados")
    );

    snapshot.forEach((documento) => {

        ingresados.push({
            id: documento.id,
            nombre: documento.data().nombre
        });

    });

}


// ===============================
// MOSTRAR TABLA
// ===============================

function mostrarInvitados(lista) {

    const tabla =
        document.getElementById(
            "tablaInvitados"
        );

    tabla.innerHTML = "";

    lista.forEach((persona) => {

        const nombre =
            persona["Nombre Completo"] || "";

        const ingresado =
            ingresados.find(
                item => item.nombre === nombre
            );

        const fila =
            document.createElement("tr");

        fila.innerHTML = `

            <td>${persona["Grado"] || ""}</td>

            <td>${nombre}</td>

            <td>${persona["Sector"] || ""}</td>

            <td>${persona["Asiento"] || ""}</td>

            <td>

                ${
                    ingresado

                    ?

                    `<button
                        class="btn btn-danger"
                        onclick="deshacerIngreso('${ingresado.id}')">
                        Deshacer
                    </button>`

                    :

                    `<button
                        class="btn btn-success"
                        onclick="marcarIngreso('${nombre}')">
                        Marcar ingreso
                    </button>`
                }

            </td>

        `;

        tabla.appendChild(fila);

    });

}


// ===============================
// MARCAR INGRESO
// ===============================

async function marcarIngreso(nombre) {

    await addDoc(
        collection(window.db, "ingresados"),
        {
            nombre: nombre
        }
    );

    await cargarIngresados();

    mostrarInvitados(invitados);

    actualizarContadores();

}


// ===============================
// DESHACER
// ===============================

async function deshacerIngreso(id) {

    await deleteDoc(
        doc(window.db, "ingresados", id)
    );

    await cargarIngresados();

    mostrarInvitados(invitados);

    actualizarContadores();

}


// ===============================
// CONTADORES
// ===============================

function actualizarContadores() {

    const total =
        invitados.length;

    const totalIngresados =
        ingresados.length;

    const pendientes =
        total - totalIngresados;

    document.getElementById(
        "totalInvitados"
    ).textContent = total;

    document.getElementById(
        "totalIngresados"
    ).textContent = totalIngresados;

    document.getElementById(
        "totalPendientes"
    ).textContent = pendientes;

}


// ===============================
// BUSCADOR
// ===============================

document
.getElementById("buscador")
.addEventListener("input", function () {

    const texto =
        this.value.toLowerCase();

    const filtrados =
        invitados.filter(persona => {

            return Object.values(persona)
                .join(" ")
                .toLowerCase()
                .includes(texto);

        });

    mostrarInvitados(filtrados);

});


// ===============================
// HACER GLOBAL
// ===============================

window.marcarIngreso =
    marcarIngreso;

window.deshacerIngreso =
    deshacerIngreso;