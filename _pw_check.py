import sys
from playwright.sync_api import sync_playwright

errors = []

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page(viewport={"width": 1440, "height": 900})
    page.on("console", lambda msg: errors.append(msg.text) if msg.type == "error" else None)
    page.on("pageerror", lambda exc: errors.append(str(exc)))

    page.goto("http://127.0.0.1:5000/", wait_until="networkidle")
    page.wait_for_timeout(800)
    page.screenshot(path="_pw_full.png", full_page=False)

    # Confirmar elementos clave del nuevo layout
    checks = {
        "rail": page.locator(".rail").count(),
        "app-header": page.locator(".app-header").count(),
        "panel-slides": page.locator(".panel-slides").count(),
        "panel-canvas": page.locator(".panel-canvas").count(),
        "panel-properties": page.locator(".panel-properties").count(),
        "canvas-slide": page.locator(".canvas-slide").count(),
        "canvas-el": page.locator(".canvas-el").count(),
        "slide-list-item": page.locator(".slide-list-item").count(),
        "properties-header": page.locator(".properties-header").count(),
    }
    print("CHECKS:", checks)

    # Click en el segundo slide de la lista, confirmar que cambia el panel de propiedades
    items = page.locator(".slide-list-item")
    if items.count() >= 2:
        before = page.locator(".properties-type-name").inner_text()
        items.nth(1).click()
        page.wait_for_timeout(300)
        after = page.locator(".properties-type-name").inner_text()
        print("FOCUS CHANGE:", before, "->", after)

    # Escribir en un input de texto y confirmar que el canvas se actualiza sin perder foco
    text_inputs = page.locator('#slidesList input[type="text"]')
    if text_inputs.count() > 0:
        first_input = text_inputs.first
        first_input.click()
        first_input.fill("Texto de prueba Playwright")
        page.wait_for_timeout(300)
        still_focused = page.evaluate("document.activeElement === document.querySelector('#slidesList input[type=\"text\"]')")
        print("INPUT STILL FOCUSED AFTER TYPE:", still_focused)
        canvas_text = page.locator(".canvas-el--text").all_inner_texts()
        print("CANVAS CONTAINS TYPED TEXT:", any("prueba Playwright" in t for t in canvas_text))

    # Zoom
    page.locator("#zoomInBtn").click()
    page.wait_for_timeout(150)
    zoom_label = page.locator("#zoomLabel").inner_text()
    print("ZOOM LABEL AFTER +:", zoom_label)

    # Duplicar slide
    dup_btns = page.locator('[data-action="duplicate-slide"]')
    count_before = page.locator(".slide-list-item").count()
    if dup_btns.count() > 0:
        page.locator(".slide-list-item").first.hover()
        dup_btns.first.click(force=True)
        page.wait_for_timeout(300)
        count_after = page.locator(".slide-list-item").count()
        print("SLIDE COUNT AFTER DUPLICATE:", count_before, "->", count_after)

    page.screenshot(path="_pw_full2.png", full_page=False)
    browser.close()

print("CONSOLE/PAGE ERRORS:", errors)
if errors:
    sys.exit(1)
