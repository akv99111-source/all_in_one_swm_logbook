// Example backend logic
const isProd = process.env.VITE_CASHFREE_MODE === 'production';
const cashfreeUrl = isProd 
  ? 'https://api.cashfree.com/pg/orders' 
  : 'https://sandbox.cashfree.com/pg/orders';

const response = await fetch(cashfreeUrl, {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'x-client-id': process.env.CASHFREE_APP_ID,
    'x-client-secret': process.env.CASHFREE_SECRET_KEY, // Ensure this matches your fixed variable!
    'x-api-version': '2023-08-01'
  },
  body: JSON.stringify(orderPayload)
});
