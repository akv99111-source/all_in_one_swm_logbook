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
  
  // Choose URL based on mode
  const baseUrl = cashfreeMode === 'production' 
    ? 'https://api.cashfree.com/pg/orders' 
    : 'https://sandbox.cashfree.com/pg/orders';

  try {
    const orderPayload = {
      order_amount: amount,
      order_currency: 'INR',
      customer_details: {
        customer_id: `cust_${Date.now()}`,
        customer_name: customerName || 'SWM User',
        customer_phone: customerPhone
      },
      order_meta: {
        // You can leave this blank or point it to your frontend URL
        return_url: "https://your-frontend-url.com/" 
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
      throw new Error(data.message || 'Failed to create Cashfree order');
    }

    // Return the payment_session_id back to App.jsx
    res.status(200).json(data);
  } catch (error) {
    console.error('Cashfree Order Error:', error);
    res.status(500).json({ message: error.message });
  }
}