"""Real-browser test of the built site (needs Chrome + Playwright and a fresh `npm run build`)."""
from __future__ import annotations

import asyncio
import functools
import http.server
import os
import shutil
import threading

import pytest

from conftest import WEB, filled, report_rows

DIST = os.path.join(WEB, "dist")
SAMPLES = os.path.join(WEB, "sample_data")
pytestmark = pytest.mark.skipif(not (shutil.which("google-chrome") and os.path.exists(os.path.join(DIST, "index.html"))),
                                reason="needs Chrome and a built dist/ folder")


@pytest.fixture(scope="module")
def site():
    handler = functools.partial(http.server.SimpleHTTPRequestHandler, directory=DIST)
    handler.log_message = lambda *a: None
    srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    yield f"http://127.0.0.1:{srv.server_address[1]}/"
    srv.shutdown()


def test_full_flow_in_browser(site, tmp_path):
    from e2e_flow import flow

    a = os.path.join(SAMPLES, "Consumers_A_before.xlsx")
    b = os.path.join(SAMPLES, "Consumers_B_after.xlsx")
    paths, errors, foreign = asyncio.run(flow(site, a, b, str(tmp_path)))
    assert not errors, errors
    assert not foreign, f"the page contacted other sites: {foreign}"
    hl, rep = paths
    cells = {r["B: Cell"] for r in report_rows(rep)}
    assert filled(hl)["UpdatedConsumerData"] == cells and len(cells) == 17
    assert filled(hl)["Notes"] == set()
