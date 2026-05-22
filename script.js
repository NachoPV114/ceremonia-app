let invitados = [];

let ingresados = JSON.parse(
    localStorage.getItem('ingresados')
) || [];

// Cargar Excel automáticamente

window.onload = async function() {

    const respuesta = await fetch('Invitados.xlsx');

    const data = await respuesta.arrayBuffer();

    const workbook = XLSX.read(data, {
        type: 'array'
    });

    const hoja =
        workbook.Sheets[workbook.SheetNames[0]];

    invitados =
        XLSX.utils.sheet_to_json(hoja);

    mostrarInvitados(invitados);

    actualizarEstadisticas();

};

// Mostrar invitados

function mostrarInvitados(lista) {

    const tabla =
        document.getElementById('tablaInvitados');

    tabla.innerHTML = '';

    lista.forEach((persona) => {

        const nombreCompleto = `
            ${persona.Nombre || ''}
            ${persona['Apellido P'] || ''}
            ${persona['Apellido M'] || ''}
        `;

        const presente =
            ingresados.includes(nombreCompleto);

        const fila =
            document.createElement('tr');

        fila.innerHTML = `

            <td>
                ${persona['Cargo/Grado'] || ''}
            </td>

            <td>
                ${nombreCompleto}
            </td>

            <td>
                ${persona['Sector'] || ''}
            </td>

            <td>
                ${persona['Asiento'] || ''}
            </td>

            <td>

                <button
                    class="btn ${
                        presente
                        ? 'btn-secondary'
                        : 'btn-success'
                    } btn-presente"

                    onclick="marcarIngreso(
                        '${nombreCompleto}'
                    )"
                >

                    ${
                        presente
                        ? 'Presente'
                        : 'Ingresó'
                    }

                </button>

            </td>
        `;

        tabla.appendChild(fila);

    });

}

// Marcar ingreso

function marcarIngreso(nombre) {

    if (ingresados.includes(nombre)) {

        ingresados = ingresados.filter(
            persona => persona !== nombre
        );

    } else {

        ingresados.push(nombre);

    }

    localStorage.setItem(
        'ingresados',
        JSON.stringify(ingresados)
    );

    mostrarInvitados(invitados);

    actualizarEstadisticas();

}

// Estadísticas

function actualizarEstadisticas() {

    document.getElementById('totalInvitados')
        .innerText = invitados.length;

    document.getElementById('contadorIngresados')
        .innerText = ingresados.length;

    document.getElementById('contadorPendientes')
        .innerText =
            invitados.length - ingresados.length;

}

// Buscar invitados

document
    .getElementById('searchInput')
    .addEventListener('keyup', function() {

        const texto =
            this.value.toLowerCase();

        const filtrados =
            invitados.filter(persona => {

                const nombreCompleto = `
                    ${persona.Nombre || ''}
                    ${persona['Apellido P'] || ''}
                    ${persona['Apellido M'] || ''}
                    ${persona['Cargo/Grado'] || ''}
                `;

                return nombreCompleto
                    .toLowerCase()
                    .includes(texto);

            });

        mostrarInvitados(filtrados);

    });