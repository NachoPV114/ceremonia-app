let invitados = [];

let ingresados = JSON.parse(
    localStorage.getItem('ingresados')
) || [];

// Cargar invitados automáticamente

window.onload = async function() {

    const respuesta = await fetch('invitados.json');

    invitados = await respuesta.json();

    mostrarInvitados(invitados);

    actualizarEstadisticas();

};

// Mostrar invitados

function mostrarInvitados(lista) {

    const tabla =
        document.getElementById('tablaInvitados');

    tabla.innerHTML = '';

    lista.forEach((persona, index) => {

        const fila = document.createElement('tr');

        const presente =
            ingresados.includes(persona.nombre);

        fila.innerHTML = `

            <td>${persona.grado}</td>

            <td>${persona.nombre}</td>

            <td>${persona.sector}</td>

            <td>${persona.asiento}</td>

            <td>

                <button
                    class="btn ${
                        presente
                        ? 'btn-secondary'
                        : 'btn-success'
                    } btn-presente"

                    onclick="marcarIngreso('${persona.nombre}')"
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

        // Quitar ingreso

        ingresados = ingresados.filter(
            persona => persona !== nombre
        );

    } else {

        // Agregar ingreso

        ingresados.push(nombre);

    }

    // Guardar automáticamente

    localStorage.setItem(
        'ingresados',
        JSON.stringify(ingresados)
    );

    // Actualizar pantalla

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

        const filtrados = invitados.filter(persona =>

            persona.nombre
                .toLowerCase()
                .includes(texto)

            ||

            persona.grado
                .toLowerCase()
                .includes(texto)

        );

        mostrarInvitados(filtrados);

    });