# Trabajo_pweb2
Solucion de un problema en Google Apps Script

Problema:
Imaginemos que tienes una bandeja de entrada de correo electrónico que recibe solicitudes de eventos con un formato predefinido en el asunto y contenido del correo. Por ejemplo, los correos con el asunto "Eventos" contienen detalles como el título del evento, la fecha, la hora y el lugar en el cuerpo del correo. Quieres automatizar el proceso de crear eventos en un calendario a partir de estos correos electrónicos, para evitar tener que hacerlo manualmente.

Solución:
Este código proporciona una solución automatizada para este problema. Escanea la bandeja de entrada en busca de correos electrónicos que coincidan con el asunto especificado (en este caso, "Eventos"). Una vez que encuentra estos correos electrónicos, extrae la información relevante del cuerpo del correo, como el título del evento, la fecha, la hora y el lugar. Luego, utiliza esta información para crear eventos en un calendario de Google. Además, evita la creación de eventos duplicados al mantener un registro de los ID de mensaje ya procesados.


