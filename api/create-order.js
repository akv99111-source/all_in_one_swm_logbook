export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ message: 'Method Not Allowed' });

  try {
    // 1. Get the data sent from your frontend React app
    const { amount, customerName, customerPhone } = req.body;
    
    // 2. Get your securely stored Vercel environment variables
    const appId = process.env.CASHFREE_APP_ID;
    const secretKey = process.env.CASHFREE_SECRET_KEY; 
    
    // 3. Determine if we are using the Test URL or the Live URL
    const isProd = process.env.VITE_CASHFREE_MODE === 'production';
    const url = isProd ? 'https://api.cashfree.com/pg/orders' : 'https://sandbox.cashfree.com/pg/orders';

    // 4. THIS FIXES THE ERROR: Define the order payload properly!
    const orderPayload = {
      order_amount: amount,
      order_currency: "INR",
      customer_details: {
        customer_id: `cust_${Date.now()}`,
        customer_phone: customerPhone,
        customer_name: customerName || "SWM User"
      }
    };

    // 5. Send the request to Cashfree
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-client-id': appId,
        'x-client-secret': secretKey,
        'x-api-version': '2023-08-01'
      },
      body: JSON.stringify(orderPayload)
    });

    const data = await response.json();
    
    // 6. Handle Cashfree API rejections (like bad keys) gracefully
    if (!response.ok) {
      return res.status(response.status).json({ message: data.message || "Cashfree API Error" });
    }

    // 7. Success! Send the session ID back to the frontend
    res.status(200).json(data);

  } catch (error) {
    console.error("Backend Error:", error);
    res.status(500).json({ message: error.message }); 
  }
}
