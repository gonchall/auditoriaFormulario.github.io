document.addEventListener("DOMContentLoaded", function () {
    console.log("Script cargado correctamente");

    const campoCorreoEmpresa = document.getElementById("cra3d_correoempresa14fa");
    const campoEmpresaAuditora = document.getElementById("cra3d_nombreempresaauditora14fa");
    const campoNombreAuditor = document.getElementById("cra3d_nombreauditorresponsable14fa");
    const campoCorreoAuditor = document.getElementById("cra3d_correoauditorresponsable14fa");
    const campoFechaInicio = document.getElementById("cra3d_fechainicioauditoria14fa"); // ← actualiza si el ID es distinto
    const campoFechaInforme = document.getElementById("cra3d_fechaestimadainformeauditoria14fa"); // ← actualiza si el ID es distinto

    const URL_FLUJO_AUTOCOMPLETAR = "https://default7042d393b8c4451aa8f811f6aed5da.85.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/75ddb566fde0445c8caddce895843c48/triggers/manual/paths/invoke/?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Y0vVfcdWjgMVGq0PwINaRMuhZ7U0-ehdPkGj_To_PyI";

    const URL_FLUJO_ACTUALIZAR = "https://default7042d393b8c4451aa8f811f6aed5da.85.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/3260704b65884a30af1b8d8602a0db7a/triggers/manual/paths/invoke/?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=E6VyfXl22cZQJLQWaPm-GyKot4Q-LwKq0g5zxxrqbfk"; 

    if (!campoCorreoEmpresa) {
        console.warn("No se encontró el campo Correo Empresa");
        return;
    }


    // Autocompletar al salir del campo correo
    campoCorreoEmpresa.addEventListener("blur", function () {
        const correo = campoCorreoEmpresa.value.trim();
        console.log("Correo capturado:", correo);
        if (!correo) return;

        fetch(URL_FLUJO_AUTOCOMPLETAR, {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({ correo_empresa: correo })
        })
        .then(res => {
            console.log("Respuesta del flujo:", res.status);
            if (!res.ok) throw new Error("Error en la respuesta del flujo");
            return res.json();
        })
        .then(data => {
            console.log("Datos recibidos:", data);

            if (campoEmpresaAuditora) {
                campoEmpresaAuditora.value = data.empresa_auditora || "";
                campoEmpresaAuditora.setAttribute("style", "color: #888;");
                campoEmpresaAuditora.addEventListener("input", function () {
                    this.setAttribute("style", "color: #000;");
                });
            }

            if (campoNombreAuditor) {
                campoNombreAuditor.value = data.nombre_auditor || "";
                campoNombreAuditor.setAttribute("style", "color: #888;");
                campoNombreAuditor.addEventListener("input", function () {
                    this.setAttribute("style", "color: #000;");
                });
            }

            if (campoCorreoAuditor) {
                campoCorreoAuditor.value = data.correo_auditor || "";
                campoCorreoAuditor.setAttribute("style", "color: #888;");
                campoCorreoAuditor.addEventListener("input", function () {
                    this.setAttribute("style", "color: #000;");
                });
            }
        })
        .catch(error => {
            console.error("Error al autocompletar:", error);
        });
    });

    // Actualizar Excel al enviar el formulario
const botonEnviar = document.getElementById("InsertButton");

    if (botonEnviar) {
        botonEnviar.addEventListener("click", function () {
            console.log("Botón Enviar clicado → actualizando Excel");

            fetch(URL_FLUJO_ACTUALIZAR, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    correo_empresa: campoCorreoEmpresa.value,
                    nombre_empresa_auditora: campoEmpresaAuditora.value,
                    nombre_auditor: campoNombreAuditor.value,
                    correo_auditor: campoCorreoAuditor.value,
                    fecha_inicio_auditoria: campoFechaInicio ? campoFechaInicio.value : "",
                    fecha_estimacion_informe: campoFechaInforme ? campoFechaInforme.value : "",
                    respuesta_formulario: "RESPONDIDO",
                    estado_auditoria: "2"
                })
            })
            .then(res => {
                console.log("Respuesta del flujo:", res.status);
                return res.json();
            })
            .then(data => {
                console.log("Respuesta de Power Automate:", data);
            })
            .catch(error => {
                console.error("Error al actualizar Excel:", error);
            });
        });
    } else {
        console.warn("No se encontró el botón de envío.");
    }

});
