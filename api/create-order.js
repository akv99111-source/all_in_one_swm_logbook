export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ message: 'Method Not Allowed' });
  }

  try {
    const { amount, customerName, customerPhone } = req.body;
    
    const appId = process.env.CASHFREE_APP_ID;
    const secretKey = process.env.CASHFREE_SECRET_KEY;
    const mode = process.env.VITE_CASHFREE_MODE || 'production';

    // Diagnostic check to catch missing keys instantly
    if (!appId || !secretKey) {
      console.error("Missing Cashfree Credentials:", { hasAppId: !!appId, hasSecret: !!secretKey });
      return res.status(500).json({ message: 'Server configuration error: Cashfree API keys missing in environment variables.' });
    }

    const url = mode === 'production' 
      ? 'https://api.cashfree.com/pg/orders' 
      : 'https://sandbox.cashfree.com/pg/orders';

    const orderPayload = {
      order_amount: amount,
      order_currency: "INR",
      customer_details: {
        customer_id: `cust_${Date.now()}`,
        customer_phone: customerPhone || "9999999999",
        customer_name: customerName || "SWM User"
      }
    };

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-client-id': appId.trim(),
        'x-client-secret': secretKey.trim(),
        'x-api-version': '2023-08-01'
      },
      body: JSON.stringify(orderPayload)
    });

    const data = await response.json();
    
    if (!response.ok) {
      console.error("Cashfree API Rejection:", data);
      return res.status(response.status).json({ message: data.message || "Cashfree Authentication Failed" });
    }

    return res.status(200).json(data);

  } catch (error) {
    console.error("Backend Execution Error:", error);
    return res.status(500).json({ message: error.message });
  }
}
