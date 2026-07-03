// Cloudflare Worker for pdffreely.com.
//
// Static assets are served directly by the Workers assets layer; this script
// only receives requests that match no asset. Its job is to serve the large
// LibreOffice WASM engine files (loaded on demand by the converter at / and
// the organizer at /organize/) from R2, same-origin, so no cross-origin
// CORS/CORP setup is needed under the pages' COEP policy.
//
// URL scheme:  /wasm/<version>/<file>   →  R2 object  libreoffice/<version>/<file>
// A pre-compressed `<file>.br` object is preferred when the client accepts
// brotli (all modern browsers), cutting ~262 MB of engine to ~60 MB on the wire.

const TYPES = {
  '.js': 'text/javascript; charset=utf-8',
  '.wasm': 'application/wasm',
  '.data': 'application/octet-stream',
  '.metadata': 'application/octet-stream',
  '.html': 'text/html; charset=utf-8',
};

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const match = url.pathname.match(/^\/wasm\/([A-Za-z0-9._-]+)\/([A-Za-z0-9._-]+)$/);
    if (!match) return env.ASSETS.fetch(request);

    if (request.method !== 'GET' && request.method !== 'HEAD') {
      return new Response('method not allowed', { status: 405 });
    }
    const key = `libreoffice/${match[1]}/${match[2]}`;
    // The runtime normalizes the worker-visible Accept-Encoding header to
    // "br, gzip"; cf.clientAcceptEncoding preserves what the client really
    // sent. Serving the raw object to non-brotli clients avoids the runtime's
    // br→gzip transcode path entirely.
    const clientEncodings =
      request.cf?.clientAcceptEncoding ?? request.headers.get('Accept-Encoding') ?? '';
    const acceptsBr = /\bbr\b/.test(clientEncodings);

    let object = null;
    let encoding = null;
    if (acceptsBr) {
      object = await env.WASM_BUCKET.get(key + '.br');
      if (object) encoding = 'br';
    }
    if (!object) object = await env.WASM_BUCKET.get(key);
    if (!object) return new Response('not found', { status: 404 });

    const ext = match[2].slice(match[2].lastIndexOf('.'));
    // No explicit Content-Length: the runtime sets it, and may transcode the
    // brotli body for clients that do not accept br.
    const headers = new Headers({
      'Content-Type': TYPES[ext] || 'application/octet-stream',
      // The version is in the path, so these are immutable.
      'Cache-Control': 'public, max-age=31536000, immutable',
      'Cross-Origin-Resource-Policy': 'same-origin',
      'X-Content-Type-Options': 'nosniff',
      'ETag': object.httpEtag,
    });
    if (encoding) headers.set('Content-Encoding', encoding);

    return new Response(request.method === 'HEAD' ? null : object.body, {
      headers,
      // Pass pre-compressed bytes through untouched.
      encodeBody: encoding ? 'manual' : 'automatic',
    });
  },
};
