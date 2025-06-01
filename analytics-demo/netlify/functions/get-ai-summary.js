// Importa el SDK de Google Generative AI (asegúrate de instalarlo como dependencia si es necesario para un entorno local,
// pero para Netlify Functions, puedes empaquetarlo o confiar en dependencias manejadas por Netlify si configuras un package.json)
// Para simplicidad inicial, vamos a usar 'node-fetch' para hacer la llamada HTTP directamente si el SDK es complejo de configurar sin un package.json en Netlify Functions básicas.
// Usaremos el endpoint REST directamente para evitar dependencias complejas en la función serverless básica.
const fetch = require('node-fetch'); // Netlify Functions lo soporta o puedes usar el 'https` module nativo de Node.js

exports.handler = async function(event, context) {
    // 1. Obtener los datos enviados desde el frontend (app.js)
    // event.body contendrá los datos enviados como un string JSON
    if (event.httpMethod !== 'POST') {
        return { statusCode: 405, body: 'Method Not Allowed' };
    }

    let requestBody;
    try {
        requestBody = JSON.parse(event.body);
    } catch (error) {
        console.error("Error parsing request body:", error);
        return { statusCode: 400, body: JSON.stringify({ error: 'Cuerpo de la solicitud inválido.' }) };
    }

    const { chartType, xAxisColumnName, yAxisColumnName, keyDataInsights } = requestBody;

    if (!chartType || !keyDataInsights) {
        return { statusCode: 400, body: JSON.stringify({ error: 'Faltan parámetros chartType o keyDataInsights.' }) };
    }

    // 2. Obtener la Clave API de Gemini de las variables de entorno de Netlify
    const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
    if (!GEMINI_API_KEY) {
        console.error("GEMINI_API_KEY no está configurada en las variables de entorno.");
        return { statusCode: 500, body: JSON.stringify({ error: 'Error de configuración del servidor.' }) };
    }

    // 3. Construir el prompt para Gemini (ajusta según tus necesidades)
    const promptText = `Eres un consultor de Business Intelligence conciso y perspicaz. A partir del siguiente resumen de un gráfico de "${chartType}" (donde el eje X es "${xAxisColumnName || 'no especificado'}" y el eje Y es "${yAxisColumnName || 'no especificado'}", y los datos clave son: "${keyDataInsights}"):

Genera un insight adicional en MÁXIMO DOS FRASES que sea revelador y despierte la curiosidad del usuario sobre el potencial de un análisis más profundo. Evita conclusiones definitivas, en su lugar, sugiere posibles áreas de interés o preguntas que valdría la pena investigar más a fondo con estos datos. No saludes ni despidas. Sé breve y directo.`;

    const API_URL = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${GEMINI_API_KEY}`;

    try {
        const response = await fetch(API_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                contents: [{
                    parts: [{ text: promptText }]
                }],
                // Opcional: Ajustar parámetros de generación
                // generationConfig: {
                //   temperature: 0.7,
                //   maxOutputTokens: 100,
                // }
            }),
        });

        if (!response.ok) {
            const errorBody = await response.text();
            console.error(`Error de la API de Gemini: ${response.status}`, errorBody);
            return { statusCode: response.status, body: JSON.stringify({ error: `Error de la API de Gemini: ${errorBody}` }) };
        }

        const data = await response.json();

        // Extraer el texto de la respuesta de Gemini
        // La estructura de la respuesta puede variar un poco, verifica la documentación de Gemini API
        let aiText = "No se pudo generar el análisis en este momento.";
        if (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts && data.candidates[0].content.parts[0]) {
            aiText = data.candidates[0].content.parts[0].text;
        } else {
             console.warn("Respuesta de Gemini no tuvo la estructura esperada:", JSON.stringify(data, null, 2));
        }


        return {
            statusCode: 200,
            body: JSON.stringify({ aiSummary: aiText }),
        };

    } catch (error) {
        console.error("Error llamando a la API de Gemini:", error);
        return { statusCode: 500, body: JSON.stringify({ error: 'Error al contactar el servicio de IA.' }) };
    }
};
