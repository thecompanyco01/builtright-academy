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

    // Try metadata first, then fall back to line item price lookup
    let productKey = session.metadata?.product_key;
    
    if (!productKey) {
      // Map Stripe price IDs to product keys
      const priceMap = {
        'price_1T9AhyPwH6Vr5IXOZ9l7UfG9': 'invoice-pro',
        'price_1T9AhzPwH6Vr5IXOCClJ5NL7': 'pro-bundle',
        'price_1T9Ai1PwH6Vr5IXOV6vcx4eS': 'complete-kit'
      };
      // Expand line items to find the product
      try {
        const lineItems = await stripe.checkout.sessions.listLineItems(session_id, { limit: 1 });
        const priceId = lineItems?.data?.[0]?.price?.id;
        productKey = priceMap[priceId] || 'pro-bundle';
      } catch {
        productKey = 'pro-bundle';
      }
    }

    // Also accept product hint from query param
    const productHint = req.query.product;
    if (!productKey && productHint && ['invoice-pro','pro-bundle','complete-kit'].includes(productHint)) {
      productKey = productHint;
    }

    const files = {
      'invoice-pro': ['contractor-invoice.xlsx'],
      'pro-bundle': ['pro-bundle.zip'],
      'complete-kit': ['complete-kit.zip']
    };
    return res.json({ paid: true, product: productKey, files: files[productKey] || files['pro-bundle'] });
  } catch (err) {
    return res.status(500).json({ error: 'Verification failed' });
  }
}
