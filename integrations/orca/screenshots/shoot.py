import sys, asyncio, functools, http.server, threading
from playwright.async_api import async_playwright
root, deck, out, w, h = sys.argv[1], sys.argv[2], sys.argv[3], int(sys.argv[4]), int(sys.argv[5])
lang = sys.argv[6] if len(sys.argv) > 6 else 'en-US'
class Q(http.server.SimpleHTTPRequestHandler):
    extensions_map = {**http.server.SimpleHTTPRequestHandler.extensions_map, '.wasm':'application/wasm', '.js':'text/javascript'}
    def log_message(self, *a): pass
srv = http.server.ThreadingHTTPServer(('127.0.0.1',0), functools.partial(Q, directory=root))
threading.Thread(target=srv.serve_forever, daemon=True).start()
MOCK = """window.orca={document:{session: async (o)=>{const b=await (await fetch('/decks/%s')).arrayBuffer();
return {value:o.decode({kind:'bytes',bytes:b}),readOnly:false,onReplace(){},change(){},flush:async()=>{}};}},on(){}};""" % deck
async def main():
    async with async_playwright() as p:
        b = await p.chromium.launch(executable_path='/usr/bin/google-chrome')
        ctx = await b.new_context(viewport={'width':1100,'height':688}, locale=lang, device_scale_factor=w/1100)
        pg = await ctx.new_page()
        await pg.add_init_script(MOCK)
        await pg.goto(f'http://127.0.0.1:{srv.server_port}/plugin/editor/index.html')
        await pg.wait_for_function("document.querySelectorAll('.slide-thumb').length > 0", timeout=90000)
        await pg.wait_for_timeout(2500)
        await pg.screenshot(path=out)
        await b.close()
asyncio.run(main())
