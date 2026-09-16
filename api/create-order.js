export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ message: 'Method Not Allowed' });

  try {
    const { amount, customerName, customerPhone } = req.body;
    
    const appId = process.env.CASHFREE_APP_ID;
    const secretKey = process.env.CASHFREE_SECRET_KEY; 
    
    const isProd = process.env.VITE_CASHFREE_MODE === 'production';
    const url = isProd ? 'https://api.cashfree.com/pg/orders' : 'https://sandbox.cashfree.com/pg/orders';

    // THIS IS THE MISSING PIECE THAT CAUSES THE REFERENCE ERROR!
    const orderPayload = {
      order_amount: amount,
      order_currency: "INR",
      customer_details: {
        customer_id: `cust_${Date.now()}`,
        customer_phone: customerPhone,
        customer_name: customerName || "SWM User"
      }
    };

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-client-id': appId,
        'x-client-secret': secretKey,
        'x-api-version': '2023-08-01'
      },
      body: JSON.stringify(orderPayload) // It was crashing here because it didn't know what orderPayload was
    });

    const data = await response.json();
    
    if (!response.ok) {
      return res.status(response.status).json({ message: data.message || "Cashfree API Error" });
    }

    res.status(200).json(data);

  } catch (error) {
    console.error("Backend Error:", error);
    res.status(500).json({ message: error.message }); 
  }
}
