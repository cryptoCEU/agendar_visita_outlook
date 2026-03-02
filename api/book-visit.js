const msal = require('@azure/msal-node');
const axios = require('axios');

const config = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  }
};

const cca = new msal.ConfidentialClientApplication(config);

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Solo POST permitido' });
  }

  const { start, end, user_name, user_email, notes = '' } = req.body;

  if (!start || !end || !user_name || !user_email) {
    return res.status(400).json({ 
      status: 'error', 
      message: 'Faltan: start, end, user_name, user_email' 
    });
  }

  try {
    // 1. Token
    const tokenResponse = await cca.acquireTokenByClientCredential({
      scopes: ['https://graph.microsoft.com/.default']
    });
    const token = tokenResponse.accessToken;
    const headers = { Authorization: `Bearer ${token}` };

    // 2. Check disponibilidad
    const checkUrl = `https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${start}&endDateTime=${end}`;
    const checkRes = await axios.get(checkUrl, { headers });
    
    if (checkRes.data.value.length > 0) {
      return res.json({ 
        status: 'busy', 
        message: 'Franja ocupada. ¿Otra hora?' 
      });
    }

    // 3. Crear evento
    const eventBody = {
      subject: `🧭 Visita: ${user_name}`,
      body: { 
        contentType: 'HTML', 
        content: `<p><strong>${user_name}</strong> (${user_email})<br>Notas: ${notes}</p>`
      },
      start: { dateTime: start, timeZone: 'Europe/Madrid' },
      end: { dateTime: end, timeZone: 'Europe/Madrid' },
      location: { displayName: 'Oficina Madrid' }
    };

    const createRes = await axios.post('https://graph.microsoft.com/v1.0/me/events', eventBody, { 
      headers: { ...headers, 'Content-Type': 'application/json' } 
    });

    res.json({ 
      status: 'booked', 
      event_id: createRes.data.id,
      start: createRes.data.start.dateTime,
      end: createRes.data.end.dateTime
    });

  } catch (error) {
    console.error('Error:', error.message);
    res.status(500).json({ status: 'error', message: error.message });
  }
};
