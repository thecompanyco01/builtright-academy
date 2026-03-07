export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });
  
  try {
    const { email, source, lead_magnet } = req.body;
    
    if (!email || !email.includes('@') || !email.includes('.')) {
      return res.status(400).json({ error: 'Valid email required' });
    }

    // Log to Vercel logs for now — structured for easy parsing
    // Format: LEAD|email|source|lead_magnet|timestamp
    const entry = {
      type: 'LEAD',
      email: email.toLowerCase().trim(),
      source: source || 'unknown',
      lead_magnet: lead_magnet || 'none',
      timestamp: new Date().toISOString()
    };
    
    console.log(`LEAD|${entry.email}|${entry.source}|${entry.lead_magnet}|${entry.timestamp}`);
    
    return res.status(200).json({ 
      success: true, 
      message: 'You\'re in! Check your email shortly.' 
    });
  } catch (err) {
    console.error('Subscribe error:', err);
    return res.status(500).json({ error: 'Server error' });
  }
}
