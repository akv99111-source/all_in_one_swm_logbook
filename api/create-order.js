export default async function handler(req, res) {
  // CORS & Method Check
  if (req.method !== 'POST') {
    return res.status(405).json({ message: 'Only POST requests allowed' });
  }

  const { amount, customerName, customerPhone } = req.body;

  // Environment Variables
  const appId = process.env.CASHFREE_APP_ID;
  const secretKey = process.env.CASHFREE_SECRET_KEY;
  const cashfreeMode = process.env.VITE_CASHFREE_MODE || 'sandbox';

  if (!appId || !secretKey) {
    return res.status(500).json({ message: 'Missing Cashfree API keys in Vercel environment variables.' });
  }

  // Choose URL based on mode
  const baseUrl = cashfreeMode === 'production' 
    ? 'https://api.cashfree.com/pg/orders' 
    : 'https://sandbox.cashfree.com/pg/orders';

  // Dynamic Host Detection for Return URL
  const protocol = req.headers['x-forwarded-proto'] || 'https';
  const host = req.headers['host'];
  const returnUrl = `${protocol}://${host}/`;

  try {
    const orderPayload = {
      order_id: `order_${Date.now()}_${Math.floor(Math.random() * 1000)}`,
      order_amount: Number(amount),
      order_currency: 'INR',
      customer_details: {
        customer_id: `cust_${Date.now()}`,
        customer_name: customerName || 'SWM User',
        customer_phone: customerPhone
      },
      order_meta: {
        return_url: returnUrl
      }
    };

    const response = await fetch(baseUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-version': '2023-08-01',
        'x-client-id': appId,
        'x-client-secret': secretKey
      },
      body: JSON.stringify(orderPayload)
    });

    const data = await response.json();

    if (!response.ok) {
      return res.status(response.status).json({ 
        message: data.message || 'Failed to create Cashfree order',
        details: data 
      });
    }

    // Return order object containing payment_session_id to App.jsx
    return res.status(200).json(data);
  } catch (error) {
    console.error('Cashfree Order Error:', error);
    return res.status(500).json({ message: error.message });
  }
}
