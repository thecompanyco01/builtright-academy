export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const stripe = require('stripe')(process.env.AGENT4_STRIPE_KEY);
  const { session_id } = req.query;
  if (!session_id) return res.status(400).json({ error: 'Missing session_id' });

  try {
    const session = await stripe.checkout.sessions.retrieve(session_id);
    if (session.payment_status !== 'paid') return res.json({ paid: false });

    const productKey = session.metadata?.product_key || 'pro-bundle';
    const files = {
      'invoice-pro': ['contractor-invoice.xlsx'],
      'pro-bundle': ['pro-bundle.zip'],
      'complete-kit': ['complete-kit.zip']
    };
    return res.json({ paid: true, product: productKey, files: files[productKey] || [] });
  } catch (err) {
    return res.status(500).json({ error: 'Verification failed' });
  }
}
