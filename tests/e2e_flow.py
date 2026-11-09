"""Drive the built web app in Chrome: upload, run, download both files. Used by test_e2e.py and for screenshots."""
import asyncio
import os
import sys

from playwright.async_api import async_playwright

SD = "/home/claude/Excel_CSV_Comparator/sample_data/"


async def flow(url, a, b, outdir, shots=None, template=None, header_a=None):
    async with async_playwright() as p:
        br = await p.chromium.launch(executable_path="/usr/bin/google-chrome", args=["--no-sandbox"])
        ctx = await br.new_context(viewport={"width": 1280, "height": 900}, accept_downloads=True)
        pg = await ctx.new_page()
        errors = []
        pg.on("pageerror", lambda e: errors.append(str(e)))
        pg.on("console", lambda m: errors.append(m.text) if m.type == "error" else None)
        reqs = []
        pg.on("request", lambda r: reqs.append(r.url))

        async def shot(name):
            if shots:
                await pg.screenshot(path=os.path.join(shots, name), full_page=True)

        await pg.goto(url)
        await pg.wait_for_timeout(1500)
        await shot("1_files_empty.png")
        await pg.locator('input[data-side="A"]').set_input_files(a)
        await pg.wait_for_selector("text=Ready")
        await pg.locator('input[data-side="B"]').set_input_files(b)
        await pg.locator('.tag.ok').nth(1).wait_for()
        await shot("1_files.png")
        nxt = pg.get_by_role("button", name="Next →")
        await nxt.click(); await pg.wait_for_timeout(800)
        if header_a:
            await pg.locator('input[type=number]').first.fill(str(header_a)); await pg.wait_for_timeout(500)
        await shot("2_sheets.png")
        await nxt.click(); await pg.wait_for_timeout(800)
        await shot("3_match.png")
        await nxt.click(); await pg.wait_for_timeout(500)
        await pg.get_by_text("Upper / lower case").click()
        await pg.get_by_text("Spaces at the start or end").click()
        await pg.get_by_text("Number format (1000").click()
        await pg.wait_for_timeout(400)
        await shot("4_rules.png")
        await nxt.click(); await pg.wait_for_timeout(500)
        await pg.wait_for_selector("button:has-text('Start comparison'):not([disabled])", timeout=60000)
        await shot("5_compare.png")
        await pg.get_by_role("button", name="Start comparison").click()
        try:
            await pg.wait_for_selector("button:has-text('Run again')", timeout=int(os.environ.get("E2E_WAIT", "300000")))
        except Exception:
            await pg.screenshot(path="/tmp/e2e_fail.png", full_page=True)
            print("PAGE ERRORS:", errors[-10:])
            raise
        await shot("5_compare_done.png")
        await pg.get_by_role("button", name="See results").click(); await pg.wait_for_timeout(800)
        await shot("6_results.png")
        await pg.get_by_role("button", name="Changed Cells").click(); await pg.wait_for_timeout(1200)
        await shot("6_results_cells.png")
        await pg.get_by_role("button", name="Validation").click(); await pg.wait_for_timeout(500)
        await shot("6_results_validation.png")
        await pg.get_by_role("button", name="Go to downloads").click(); await pg.wait_for_timeout(500)
        await shot("7_download.png")
        paths = []
        for label in ("Download highlighted copy", "Download report"):
            async with pg.expect_download() as dl:
                await pg.get_by_role("button", name=label).click()
            d = await dl.value
            path = os.path.join(outdir, d.suggested_filename)
            await d.save_as(path)
            paths.append(path)
        if shots:
            await pg.set_viewport_size({"width": 400, "height": 820}); await pg.wait_for_timeout(500)
            await pg.screenshot(path=os.path.join(shots, "8_mobile.png"))
        origin = url.split("/")[2]
        foreign = [r for r in reqs if not r.startswith(("data:", "blob:")) and r.split("/")[2] != origin]
        await br.close()
        return paths, errors, foreign


if __name__ == "__main__":
    a, b, out, shots = sys.argv[1:5]
    print(asyncio.run(flow("http://localhost:4173/", a, b, out, shots)))
