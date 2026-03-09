import { readFileSync } from 'fs';
import { join } from 'path';

export default async function handler(req, res) {
  const stripe = require('stripe')(process.env.AGENT4_STRIPE_KEY);
  const { file, session_id } = req.query;
  if (!file || !session_id) return res.status(400).json({ error: 'Missing params' });

  // Verify payment
  try {
    const session = await stripe.checkout.sessions.retrieve(session_id);
    if (session.payment_status !== 'paid') return res.status(403).json({ error: 'Payment not verified' });
  } catch { return res.status(403).json({ error: 'Invalid session' }); }

  // Sanitize filename
  const safeName = file.replace(/[^a-zA-Z0-9._-]/g, '');
  const allowed = ['contractor-invoice.xlsx','contractor-estimate.xlsx','job-costing-tracker.xlsx','profit-loss-tracker.xlsx','client-tracker.xlsx','contractor-proposal.xlsx','change-order.xlsx','weekly-timesheet.xlsx','cash-flow-forecast.xlsx','pro-bundle.zip','complete-kit.zip'];
  if (!allowed.includes(safeName)) return res.status(404).json({ error: 'File not found' });

  try {
    const filePath = join(process.cwd(), 'downloads', safeName);
    const data = readFileSync(filePath);
    const ext = safeName.endsWith('.zip') ? 'application/zip' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    res.setHeader('Content-Type', ext);
    res.setHeader('Content-Disposition', `attachment; filename="${safeName}"`);
    return res.status(200).send(data);
  } catch { return res.status(404).json({ error: 'File not found' }); }
}
