export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const stripe = require('stripe')(process.env.AGENT4_STRIPE_KEY);
  const products = {
    'invoice-pro': { price: 'price_1T9AhyPwH6Vr5IXOZ9l7UfG9', name: 'Invoice Template Pro' },
    'pro-bundle': { price: 'price_1T9AhzPwH6Vr5IXOCClJ5NL7', name: 'Pro Contractor Bundle' },
    'complete-kit': { price: 'price_1T9Ai1PwH6Vr5IXOV6vcx4eS', name: 'Complete Business Kit' }
  };

  const { product } = req.body || {};
  if (!products[product]) return res.status(400).json({ error: 'Invalid product' });

  try {
    const session = await stripe.checkout.sessions.create({
      payment_method_types: ['card'],
      line_items: [{ price: products[product].price, quantity: 1 }],
      mode: 'payment',
      success_url: 'https://builtrighthq.com/download?session_id={CHECKOUT_SESSION_ID}',
      cancel_url: 'https://builtrighthq.com/templates',
      metadata: { product_key: product }
    });
    return res.status(200).json({ url: session.url });
  } catch (err) {
    console.error('Checkout error:', err.message, err.type, err.code);
    return res.status(500).json({ error: 'Failed to create checkout' });
  }
}
