document.addEventListener('DOMContentLoaded', function() {
    const eventosTimeline = document.querySelectorAll('.linea-de-tiempo .evento');
    const contenidoEvento = document.querySelector('.linea-de-tiempo .contenido-evento');
    const habilidadesBarras = document.querySelectorAll('.habilidades .barra-progreso');

    // Datos de la experiencia (podríamos extraer esto del HTML también)
    const experienciaData = {
        'responsable': `
            <h3>Responsable de área BI & Reporting en Dial Database Marketing “DDM”</h3>
            <span class="fecha">Octubre 2021 – Presente</span>
            <p>Liderazgo del área de BI & Reporting, implementando IA generativa para análisis de datos, modelado para dimensionamientos en licitaciones, análisis y desarrollo de encuestas (productos, perfil de clientes, políticas), y creación de dashboards estratégicos para CEO, CCHH, ADM, IT y Operaciones. Generación de informes y análisis de métricas organizacionales.</p>
        `,
        'coordinador': `
            <h3>Coordinador en Dial Database Marketing “DDM”</h3>
            <span class="fecha">Enero 2018 – Septiembre 2021</span>
            <p>Recepción, mantenimiento y actualización de bases de datos. Coordinación con IT para validación e integración de datos. Análisis y detección de clusters para optimizar la segmentación outbound. Generación de dashboards con características clave de datos para identificar oportunidades.</p>
        `,
        'reporting': `
            <h3>Reporting en Dial Database Marketing “DDM”</h3>
            <span class="fecha">Febrero 2014 – Diciembre 2017</span>
            <p>Actualización y optimización de datos para informes internos y externos. Procesamiento y disposición de información para diversas áreas. Detección y análisis del comportamiento de clientes (existentes y potenciales). Delegación y coordinación de equipo. Confección y customización de informes. Gestión de agenda de entregables.</p>
        `,
        'analista': `
            <h3>Analista de calidad Jr. en ACC GROUP</h3>
            <span class="fecha">Diciembre 2008 – Mayo 2013</span>
            <p>Auditoría, análisis y calificación de llamadas de agentes. Análisis de métricas internas e implementación de acciones para objetivos. Capacitación de agentes. Fidelización y retención de clientes.</p>
        `
    };

    // Mostrar el contenido del primer evento al cargar la página
    if (eventosTimeline.length > 0 && experienciaData[eventosTimeline[0].dataset.rol]) {
        contenidoEvento.innerHTML = experienciaData[eventosTimeline[0].dataset.rol];
        eventosTimeline[0].classList.add('activo');
    }

    // Agregar event listeners a los eventos de la línea de tiempo
    eventosTimeline.forEach(evento => {
        evento.addEventListener('click', function() {
            // Remover la clase 'activo' de todos los eventos
            eventosTimeline.forEach(e => e.classList.remove('activo'));
            // Agregar la clase 'activo' al evento clickeado
            this.classList.add('activo');
            // Mostrar el contenido correspondiente
            const rol = this.dataset.rol;
            if (experienciaData[rol]) {
                contenidoEvento.innerHTML = experienciaData[rol];
            } else {
                contenidoEvento.innerHTML = '<p>Información no disponible.</p>';
            }
        });
    });

    // Animar las barras de progreso de las habilidades
    habilidadesBarras.forEach(barra => {
        const nivel = barra.dataset.nivel;
        barra.style.setProperty('--width', nivel + '%');
        barra.innerHTML = `<span style="position: absolute; right: 5px; color: white; font-size: 0.8em;">${nivel}%</span>`;
    });

    // Definir una variable CSS para controlar el ancho de la barra (necesario para la animación)
    const styleSheet = document.createElement("style");
    styleSheet.type = "text/css";
    styleSheet.innerText = `.barra-progreso::after { width: var(--width); }`;
    document.head.appendChild(styleSheet);
});
