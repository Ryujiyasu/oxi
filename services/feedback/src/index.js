// Takes the rating a person chose to send from inside Oxi, and nothing else.
//
// The app sends only when someone has read what is about to go and pressed the
// button; this end is written on the same footing. It accepts a rating, an
// optional comment, and the version and platform needed to make sense of them.
// It does not accept, and has nowhere to put, anything about the documents the
// person was working on.
//
// Deploy:
//   npx wrangler d1 create oxi-feedback
//   npx wrangler d1 execute oxi-feedback --remote --file schema.sql
//   npx wrangler deploy
//
// Read what has come in:
//   npx wrangler d1 execute oxi-feedback --remote \
//     --command "select at, rating, version, platform, comment from feedback order by at desc limit 40"

const CORS = {
  // The app runs from a scheme of its own, so there is no origin worth naming.
  'access-control-allow-origin': '*',
  'access-control-allow-methods': 'POST, OPTIONS',
  'access-control-allow-headers': 'content-type',
  'access-control-max-age': '86400',
};

/// A day's worth of sending from one address. A person rates the app once;
/// anything past this is not a person.
const A_DAY = 5;

const said = (status, body) =>
  new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json', ...CORS },
  });

export default {
  async fetch(request, env) {
    if (request.method === 'OPTIONS') return new Response(null, { headers: CORS });
    if (request.method !== 'POST') return said(405, { error: 'post a rating' });

    // A rating is a few hundred bytes. Anything larger is not one.
    const length = Number(request.headers.get('content-length') || 0);
    if (length > 4096) return said(413, { error: 'too long' });

    let sent;
    try {
      sent = await request.json();
    } catch {
      return said(400, { error: 'not json' });
    }

    const rating = Number(sent.rating);
    if (!Number.isInteger(rating) || rating < 1 || rating > 5) {
      return said(400, { error: 'rating must be 1 to 5' });
    }
    const comment = typeof sent.comment === 'string' ? sent.comment.slice(0, 2000) : '';
    const version = typeof sent.version === 'string' ? sent.version.slice(0, 32) : '';
    const platform = typeof sent.platform === 'string' ? sent.platform.slice(0, 64) : '';

    // The address is used to hold the rate limit and is not stored: a rating is
    // worth keeping, the address it came from is not.
    const from = request.headers.get('cf-connecting-ip') || 'unknown';
    const today = new Date().toISOString().slice(0, 10);
    const key = `${today}/${from}`;
    const seen = await env.DB.prepare(
      'insert into senders (key, times) values (?1, 1) ' +
      'on conflict(key) do update set times = times + 1 returning times',
    ).bind(key).first();
    if (seen && seen.times > A_DAY) return said(429, { error: 'that is enough for today' });

    await env.DB.prepare(
      'insert into feedback (at, rating, comment, version, platform) values (?1, ?2, ?3, ?4, ?5)',
    ).bind(new Date().toISOString(), rating, comment, version, platform).run();

    return said(200, { ok: true });
  },
};
