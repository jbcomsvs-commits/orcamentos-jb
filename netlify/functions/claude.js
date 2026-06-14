// Rate limiting simples em memória (30 chamadas por minuto por IP)
const callLog = new Map();
const RATE_LIMIT = 30;
const RATE_WINDOW = 60 * 1000;

function checkRateLimit(ip) {
  const now = Date.now();
  const calls = (callLog.get(ip) || []).filter(t => now - t < RATE_WINDOW);
  if(calls.length >= RATE_LIMIT) return false;
  calls.push(now);
  callLog.set(ip, calls);
  return true;
}

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }

  const ip = event.headers['x-forwarded-for'] || event.headers['client-ip'] || 'unknown';
  if(!checkRateLimit(ip)) {
    return {
      statusCode: 429,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ error: { message: 'Muitas requisições. Aguarde um momento.' } }),
    };
  }

  const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
  if (!ANTHROPIC_API_KEY) {
    return {
      statusCode: 500,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ error: { message: 'ANTHROPIC_API_KEY não configurada no Netlify' } }),
    };
  }

  try {
    const body = JSON.parse(event.body || '{}');

    // Limitar max_tokens para reduzir tempo de resposta em PDFs grandes
    if (body.max_tokens && body.max_tokens > 4000) {
      body.max_tokens = 4000;
    }

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify(body),
    });

    const data = await response.json();

    return {
      statusCode: response.status,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data),
    };
  } catch (err) {
    const isTimeout = err.message?.includes('timeout') || err.message?.includes('network');
    return {
      statusCode: 500,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        error: {
          message: isTimeout
            ? 'Tempo limite excedido. Tente com menos itens ou use o modo manual.'
            : err.message
        }
      }),
    };
  }
};
