from playwright.sync_api import sync_playwright

errors = []

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page(viewport={"width": 1440, "height": 900})
    page.on("console", lambda msg: errors.append(msg.text) if msg.type == "error" else None)
    page.on("pageerror", lambda exc: errors.append(str(exc)))

    page.goto("http://127.0.0.1:5000/", wait_until="networkidle")
    page.wait_for_timeout(600)

    # Forzar el slide enfocado a tipo "cover" y reemplazar su canvas-layout
    # cacheado por uno sintetico CON un elemento ligado al campo "title", para
    # probar la rama de vinculacion en vivo sin depender de que los datos
    # reales del usuario tengan algun elemento canonico visible.
    page.evaluate("""
        state.focusedIdx = 0;
        state.slides[0].type = 'cover';
        state.canvasLayoutCache['cover'] = {
            title: {left: 0.8, top: 2.3, width: 11.5, height: 1.5, font_size: 40, fields: ['title']}
        };
        renderSlides();
    """)
    page.wait_for_timeout(300)

    canvas_text_before = page.locator(".canvas-el--text").all_inner_texts()
    print("CANVAS ANTES:", canvas_text_before)

    title_input = page.locator('#slidesList input[data-path="title"]')
    title_input.click()
    title_input.fill("Titulo Sincronizado En Vivo")
    page.wait_for_timeout(250)

    canvas_text_after = page.locator(".canvas-el--text").all_inner_texts()
    print("CANVAS DESPUES:", canvas_text_after)
    print("LIVE BINDING OK:", any("Titulo Sincronizado En Vivo" in t for t in canvas_text_after))

    browser.close()

print("ERRORS:", errors)
