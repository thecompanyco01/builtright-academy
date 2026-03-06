// Stripe checkout session creator
// Will be activated once STRIPE_SECRET_KEY is provided
export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const stripeKey = process.env.STRIPE_SECRET_KEY;
  if (!stripeKey) {
    return res.status(503).json({ error: 'Payments not yet configured' });
  }

  try {
    const { product, email } = req.body || {};
    
    const products = {
      'pro-template-bundle': {
        name: 'Pro Contractor Template Bundle',
        description: '12 professional templates for contractors — invoices, estimates, proposals, job costing, and more.',
        price: 2900, // $29.00 in cents
      },
      'estimating-course': {
        name: 'Profitable Estimating & Bidding Course',
        description: 'Video course + templates on pricing jobs profitably. For electricians, plumbers, HVAC, and contractors.',
        price: 14700, // $147.00 founding member price
      }
    };

    const selectedProduct = products[product];
    if (!selectedProduct) {
      return res.status(400).json({ error: 'Invalid product' });
    }

    // Dynamic import of stripe
    const stripe = (await import('stripe')).default(stripeKey);
    
    const session = await stripe.checkout.sessions.create({
      payment_method_types: ['card'],
      line_items: [{
        price_data: {
          currency: 'usd',
          product_data: {
            name: selectedProduct.name,
            description: selectedProduct.description,
          },
          unit_amount: selectedProduct.price,
        },
        quantity: 1,
      }],
      mode: 'payment',
      success_url: `${req.headers.origin || 'https://builtright-academy.vercel.app'}/thank-you?product=${product}`,
      cancel_url: `${req.headers.origin || 'https://builtright-academy.vercel.app'}/templates/`,
      customer_email: email || undefined,
    });

    return res.status(200).json({ url: session.url });
  } catch (error) {
    console.error('Stripe error:', error);
    return res.status(500).json({ error: 'Payment processing error' });
  }
}
