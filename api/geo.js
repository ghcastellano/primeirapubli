// Retorna o país do visitante usando o cabeçalho de geolocalização do Vercel.
// 1ª parte (mesmo domínio) → bloqueadores de anúncio não conseguem barrar.
export default function handler(req, res) {
  const country = (req.headers['x-vercel-ip-country'] || '').toUpperCase();
  res.setHeader('Cache-Control', 'no-store');
  res.status(200).json({ country });
}
