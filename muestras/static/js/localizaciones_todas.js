function desplegar(button, className) {
    const element = button.parentElement.querySelector("." + className);

    if (!element) return;

    if (element.style.display === "block") {
        element.style.display = "none";
        button.innerHTML = button.innerHTML.replace("▼", "▶");
    } else {
        element.style.display = "block";
        button.innerHTML = button.innerHTML.replace("▶", "▼");
    }
}

function desplegarCaja(button, cajaId) {
    const ul = button.parentElement.querySelector('ul.muestras[data-caja-id="' + cajaId + '"]');
    if (!ul) return;

    // Toggle visibilidad
    if (ul.style.display === "block") {
        ul.style.display = "none";
        button.innerHTML = button.innerHTML.replace("▼", "▶");
        return;
    }

    ul.style.display = "block";
    button.innerHTML = button.innerHTML.replace("▶", "▼");

    // Si ya se cargaron las subposiciones, no volver a cargar
    if (ul.dataset.loaded === "true") return;

    // Mostrar indicador de carga
    ul.innerHTML = '<li style="color:#888; font-style:italic; padding:6px 0;"><span class="spinner-sm"></span> Cargando subposiciones…</li>';

    fetch('/api/get_subposiciones_por_caja_tree/?caja_id=' + cajaId)
        .then(function(response) { return response.json(); })
        .then(function(data) {
            ul.innerHTML = '';
            if (!data.subposiciones || data.subposiciones.length === 0) {
                ul.innerHTML = '<li style="color:#888; font-style:italic; padding:6px 0;">Sin subposiciones</li>';
                ul.dataset.loaded = "true";
                return;
            }
            data.subposiciones.forEach(function(sub) {
                var li = document.createElement('li');
                if (!sub.vacia && sub.muestra_nom_lab) {
                    var claseEstado = '';
                    if (sub.muestra_estado === 'Disponible') claseEstado = 'icono-disponible';
                    else if (sub.muestra_estado === 'Parcialmente enviada') claseEstado = 'icono-parcial';
                    li.innerHTML =
                        '<span class="estado-icono ' + claseEstado + '">' +
                        '<svg class="icono-estado" width="14" height="14" viewBox="0 0 24 24">' +
                        '<path fill="currentColor" d="M7 2v2h1v3.586l-3.707 3.707A1 1 0 0 0 4 12v8a3 3 0 0 0 3 3h10a3 3 0 0 0 3-3v-8a1 1 0 0 0-.293-.707L16 7.586V4h1V2H7Zm3 2h4v3.414l2 2V12H8V9.414l2-2V4Zm-2 10h8v6a1 1 0 0 1-1 1H9a1 1 0 0 1-1-1v-6Z"/>' +
                        '</svg></span>' +
                        '<input type="checkbox" name="subposicion" value="' + sub.id + '" class="form-control" onchange="actualizarEstadoBotonEliminar()"> ' +
                        'Subposici\u00f3n ' + sub.numero + ': ' +
                        '<a href="/archivo/detalles_muestra/' + sub.muestra_nom_lab + '">' + sub.muestra_nom_lab + '</a>';
                } else {
                    li.innerHTML =
                        '<input type="checkbox" name="subposicion" value="' + sub.id + '" class="form-control" onchange="actualizarEstadoBotonEliminar()"> ' +
                        'Subposici\u00f3n ' + sub.numero + ': Vac\u00eda';
                }
                ul.appendChild(li);
            });
            ul.dataset.loaded = "true";
        })
        .catch(function() {
            ul.innerHTML = '<li style="color:#c00; padding:6px 0;">Error al cargar subposiciones</li>';
        });
}

document.addEventListener("DOMContentLoaded", () => {
    document.querySelectorAll(
        ".estantes, .posicion_estante, .racks, .posicion_caja_rack, .cajas, .muestras"
    ).forEach(el => el.style.display = "none");

    document.querySelectorAll(".dropbtn").forEach(btn => {
        btn.innerHTML = btn.innerHTML + " "+ "▶ ";
    });
});
