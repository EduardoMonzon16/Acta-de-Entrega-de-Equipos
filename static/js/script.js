
function addRow() {
    const tbody = document.getElementById("historialBody");
    const actualRow = document.getElementById("filaActual");

    const newRow = document.createElement("tr");
    newRow.innerHTML = `
        <td><input type="date" name="historial_inicio[]" class="form-control form-control-sm"></td>
        <td><input type="date" name="historial_fin[]" class="form-control form-control-sm"></td>
        <td><input type="text" name="historial_usuario[]" class="form-control form-control-sm" required></td>
        <td class="text-center align-middle">
            <button type="button" class="btn btn-danger btn-sm" onclick="removeRow(this)">Eliminar</button>
        </td>
    `;
    tbody.insertBefore(newRow, actualRow);
}

function removeRow(button) {
    const row = button.closest("tr");
    row.remove();
}

function addEventoRow() {
    const table = document.getElementById("eventosTable").getElementsByTagName('tbody')[0];
    const newRow = table.insertRow();

    const cellFecha = newRow.insertCell(0);
    const cellObservaciones = newRow.insertCell(1);
    const cellAcciones = newRow.insertCell(2);

    cellFecha.innerHTML = `<input type="date" name="evento_fecha[]" class="form-control form-control-sm" required>`;
    cellObservaciones.innerHTML = `<input type="text" name="evento_observaciones[]" class="form-control form-control-sm" required>`;
    cellAcciones.innerHTML = `<button type="button" class="btn btn-danger btn-sm" onclick="removeEventoRow(this)">Eliminar</button>`;
    cellAcciones.classList.add('text-center', 'align-middle');
}

function removeEventoRow(button) {
    const row = button.closest("tr");
    row.remove();
}

document.addEventListener("DOMContentLoaded", function () {
    const correoInput = document.getElementById("correo");
    const usuarioInput = document.getElementById("usuario");
    const telefonoInput = document.getElementById("telefono");
    const nombreInput = document.getElementById("nombre");
    const nombreRecibeInput = document.getElementById("nombre_recibe");
    const tipoSelect = document.getElementById("tipo");
    const equipoInput = document.getElementById("equipo");
    const marcaSelect = document.getElementById("marca");
    const marcaEquipoInput = document.getElementById("marca_equipo");
    const modeloSelect = document.getElementById("modelo");
    const modeloEquipoInput = document.getElementById("modelo_equipo");
    const serialInput = document.getElementById("serial")
    const serieEquipoInput = document.getElementById("serie_equipo")

    if (correoInput && usuarioInput) {
        correoInput.addEventListener("input", function () {
            usuarioInput.value = correoInput.value.trim();
        });
    }
    
    if (serialInput && serieEquipoInput) {
        serialInput.addEventListener("input", function () {
            serieEquipoInput.value = serialInput.value.trim();
        });
    }

    if (telefonoInput) {
        telefonoInput.addEventListener("input", function () {
            let value = this.value.replace(/\D/g, '');

            if (value.length > 9) {
                value = value.substring(0, 9);
            }

            let formatted = value.replace(/(\d{3})(\d{3})(\d{0,3})/, function (_, p1, p2, p3) {
                return [p1, p2, p3].filter(Boolean).join(' ');
            });

            this.value = formatted;
        });
    }

    if (nombreInput && nombreRecibeInput) {
        nombreInput.addEventListener("input", function () {
            nombreRecibeInput.value = nombreInput.value.trim();
        });
    }

    if (tipoSelect && equipoInput) {
        tipoSelect.addEventListener("change", function () {
            const tipoSeleccionado = this.value.trim();
            equipoInput.value = tipoSeleccionado;
        });
    }

    if (marcaSelect && marcaEquipoInput) {
        marcaSelect.addEventListener("change", function () {
            const marcaSeleccionada = this.value.trim();
            marcaEquipoInput.value = marcaSeleccionada;
        });
    }

    if (modeloSelect && modeloEquipoInput) {
        modeloSelect.addEventListener("change", function () {
            const modeloSeleccionado = this.value.trim();
            modeloEquipoInput.value = modeloSeleccionado;
        });
    }

});
