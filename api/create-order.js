export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ message: 'Method Not Allowed' });
  }

  const { amount, customerName, customerPhone } = req.body;

  try {
    // 1. Setup Cashfree Payload
    const orderData = {
      order_amount: amount,
      order_currency: 'INR',
      customer_details: {
        customer_id: `cust_${Date.now()}`,
        customer_name: customerName || 'ULB User',
        customer_phone: customerPhone || '9999999999',
      },
      order_meta: {
        // Change this if you have a specific return URL
        return_url: 'https://all-in-one-swm-logbook.vercel.app/?order_id={order_id}', 
      },
    };

    // Use production URL if mode is set, otherwise default to sandbox
    const isProd = process.env.VITE_CASHFREE_MODE === 'production';
    const endpoint = isProd 
      ? 'https://api.cashfree.com/pg/orders' 
      : 'https://sandbox.cashfree.com/pg/orders';

    // 2. Make the request to Cashfree
    const response = await fetch(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-client-id': process.env.CASHFREE_APP_ID,
        'x-client-secret': process.env.CASHFREE_SECRET_KEY,
        'x-api-version': '2023-08-01', // CRITICAL: Cashfree fails without this
      },
      body: JSON.stringify(orderData),
    });

    // 3. SAFE PARSING: Read raw text first to prevent HTML/JSON crash
    const rawText = await response.text();

    // If Cashfree returned an HTML error page (starts with <)
    if (rawText.trim().startsWith('<')) {
      console.error("CASHFREE RETURNED HTML ERROR:", rawText);
      return res.status(502).json({ 
        message: "Cashfree Gateway Error: Received HTML instead of JSON. Check Vercel Logs." 
      });
    }

    // 4. Parse the valid JSON
    const data = JSON.parse(rawText);

    if (!response.ok) {
      console.error("CASHFREE API REJECTED:", data);
      return res.status(response.status).json({ 
        message: data.message || 'Payment initiation failed at Cashfree' 
      });
    }

    // 5. Success
    return res.status(200).json(data);

  } catch (error) {
    console.error("BACKEND CRASH:", error);
    return res.status(500).json({ message: error.message });
  }
}
