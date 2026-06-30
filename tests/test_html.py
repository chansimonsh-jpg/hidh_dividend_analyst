"""
test_html.py — Structural tests for static HTML files and generate_html.py.

These tests do NOT run a browser.  They verify that the generated source code
contains the expected JavaScript/HTML constructs so that runtime behaviour
(language persistence, section IDs, nav links) is correct.
"""

import os
import importlib.util
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _read(relpath):
    with open(os.path.join(ROOT, relpath), encoding="utf-8") as f:
        return f.read()


# ── helpers ────────────────────────────────────────────────────────────────

def _load_setlang_js():
    """Return the SETLANG_JS string from generate_html.py (second definition)."""
    src = _read("generate_html.py")
    # Second SETLANG_JS assignment overrides the first at runtime.
    # Split on the marker and take the last occurrence.
    marker = 'SETLANG_JS = """'
    parts = src.split(marker)
    assert len(parts) >= 3, "Expected at least 2 SETLANG_JS definitions"
    # parts[-1] is after the last marker; close it at the triple-quote end
    closing = '"""\n'
    block = parts[-1]
    end = block.find(closing)
    assert end != -1, "Could not find closing triple-quote for SETLANG_JS"
    return block[:end]


# ══════════════════════════════════════════════════════════════════════════
# 1. generate_html.py — SETLANG_JS language persistence
# ══════════════════════════════════════════════════════════════════════════

class TestMainPageLangPersistence:
    """The main page's setLang() must save and restore language via localStorage."""

    def setup_method(self):
        self.js = _load_setlang_js()

    def test_setlang_saves_to_localstorage(self):
        """setLang() must write the chosen language to localStorage."""
        assert "localStorage.setItem('hidh_lang'" in self.js, (
            "setLang() does not save to localStorage — language will reset on navigation"
        )

    def test_setlang_init_iife_present(self):
        """An IIFE must read localStorage on page load and apply the saved language."""
        assert "localStorage.getItem('hidh_lang')" in self.js, (
            "No localStorage read on page load — language will not be restored after navigation"
        )

    def test_setlang_init_calls_setlang(self):
        """The init IIFE must actually call setLang() with the stored value."""
        js = self.js
        # The IIFE pattern: var l=localStorage.getItem(...); if(l...) setLang(l)
        assert "setLang(l)" in js or "setLang(saved)" in js, (
            "Init code reads localStorage but does not call setLang() with the result"
        )


# ══════════════════════════════════════════════════════════════════════════
# 1b. generate_html.py — market page builder
# ══════════════════════════════════════════════════════════════════════════

class TestMarketPageBuilder:
    """_build_market_html() must produce fully i18n-ready market pages."""

    def setup_method(self):
        self.src = _read("generate_html.py")

    def test_builder_function_exists(self):
        assert "_build_market_html(" in self.src

    def test_market_files_named_market_info(self):
        """Generated files must use the *_market_info.html naming convention."""
        assert "_market_info.html" in self.src
        # Old bare names must not appear as output filenames
        assert '"us.html"' not in self.src or "us_market_info" in self.src

    def test_nav_links_use_market_info_names(self):
        """Nav hrefs inside the HTML template must use *_market_info.html."""
        for mkt in ["us", "hk", "uk", "cn"]:
            assert f"/{mkt}_market_info.html" in self.src, \
                f"Nav link for {mkt}_market_info.html missing"

    def test_section_title_ids_present(self):
        """Section titles need id= for the i18n extension to target them."""
        assert 'id="watch-title"' in self.src
        assert 'id="avoid-title"' in self.src

    def test_stat_semantic_classes(self):
        """Stat strip elements need classes so MKT_I18N_EXT can translate them."""
        assert 'class="stat-strong"' in self.src
        assert 'class="stat-watch"'  in self.src
        assert 'class="stat-unit"'   in self.src

    def test_empty_state_classes(self):
        """Empty-state paragraphs need classes for translation."""
        assert 'class="no-picks"' in self.src
        assert 'class="no-avoid"' in self.src

    def test_footer_has_about_and_privacy(self):
        """Market page footer must have About and Privacy links (matching main page)."""
        tmpl = self._get_builder_template()
        assert '/about.html'   in tmpl, "footer missing about.html link"
        assert '/privacy.html' in tmpl, "footer missing privacy.html link"

    def test_mkt_i18n_ext_constant_defined(self):
        """MKT_I18N_EXT must be a module-level constant (not inside f-string)."""
        assert "MKT_I18N_EXT" in self.src

    def test_generation_loop_injects_ext(self):
        """The per-market generation loop must inject both SETLANG_JS and MKT_I18N_EXT."""
        assert "MKT_I18N_EXT}" in self.src   # f-string: {MKT_I18N_EXT}

    def test_workflow_deploys_market_pages(self):
        """HTML-deploy workflows must copy market pages to _pages_site/.
        (daily-data-update.yml only updates Excel — no HTML deploy needed there.)
        """
        for wf in ["daily-html-deploy.yml", "daily-dividend-update.yml"]:
            wf_src = _read(f".github/workflows/{wf}")
            assert "market_info.html" in wf_src, \
                f"{wf} does not deploy market_info pages"

    def test_sparkline_included_in_builder(self):
        """renderSparkline must be included in the market page template."""
        assert "renderSparkline" in self.src

    # ── Market page structure completeness ────────────────────────────────

    def _get_builder_template(self):
        """Return the full _build_market_html function body."""
        idx = self.src.find("def _build_market_html(")
        assert idx != -1, "_build_market_html not found"
        return self.src[idx:idx + 20000]

    def test_market_page_has_header(self):
        """Market pages must have a proper header."""
        tmpl = self._get_builder_template()
        assert 'class="header"'       in tmpl
        assert 'class="header-inner"' in tmpl
        assert 'class="logo"'         in tmpl

    def test_market_page_has_footer(self):
        """Market pages must have a dark footer."""
        tmpl = self._get_builder_template()
        assert 'class="footer"'       in tmpl
        assert 'class="footer-inner"' in tmpl

    def test_market_page_has_lang_switcher(self):
        """Market pages must have the 繁體/简体/English language buttons."""
        tmpl = self._get_builder_template()
        assert 'btn-zh-hk' in tmpl
        assert 'btn-zh-cn' in tmpl
        assert 'btn-en'    in tmpl
        assert 'setLang'   in tmpl

    def test_market_page_nav_back_to_index(self):
        """Market pages must have a nav link back to the main page (/)."""
        tmpl = self._get_builder_template()
        assert 'href="/"' in tmpl     # the Home link in nav
        assert '主頁'     in tmpl

    def test_market_page_has_dropdown_navigation(self):
        """Market pages must use a <select> dropdown to navigate between markets."""
        src = self.src
        assert 'id="mkt-select"' in src,           "mkt-select dropdown missing"
        assert "onchange" in src,                  "dropdown onchange handler missing"
        for mkt in ["us", "hk", "uk", "cn"]:
            assert f"/{mkt}_market_info.html" in src, \
                f"No option for /{mkt}_market_info.html in dropdown"

    def test_market_dropdown_placeholder_is_market_not_home(self):
        """Dropdown placeholder must say '市場' / 'Market', not '主頁' / 'Home'."""
        src = self.src
        assert "mktLabel:'── 市場 ──'" in src,   "zh-hk mktLabel not set to 市場"
        assert "mktLabel:'── 市场 ──'" in src,   "zh-cn mktLabel not set to 市场"
        assert "mktLabel:'── Market ──'" in src, "en mktLabel not set to Market"

    def test_market_dropdown_options_translate(self):
        """Dropdown option labels must translate for all 3 languages."""
        src = self.src
        for lang_key in ["optUS", "optHK", "optUK", "optCN"]:
            assert lang_key in src, f"{lang_key} translation missing from MKT_I18N_EXT"

    def test_index_chart_cards_are_links(self):
        """The 4 donut chart cards must be wrapped in <a> tags pointing to market pages.
        This is a true link (keyboard nav, right-click, mobile-friendly).
        """
        for mkt in ["us", "hk", "uk", "cn"]:
            assert f'<a href="/{mkt}_market_info.html"' in self.src, \
                f"<a> link missing for {mkt} chart card"

    def test_market_page_section_title_matches_main(self):
        """Market page .section-title CSS must match the main page style
        (font-size:20px, font-weight:700, color:#222).
        """
        tmpl = self._get_builder_template()
        assert "font-size:20px" in tmpl,    "Market page section-title font-size should be 20px"
        assert "font-weight:700" in tmpl,   "Market page section-title font-weight should be 700"
        assert "color:#222" in tmpl,        "Market page section-title color should be #222"


# ══════════════════════════════════════════════════════════════════════════
# 2. about.html — header, nav, persistence
# ══════════════════════════════════════════════════════════════════════════

class TestAboutHtml:
    def setup_method(self):
        self.src = _read("about.html")

    def test_has_header_with_logo(self):
        """about.html must have a proper header with the HiDH logo."""
        assert 'class="header"' in self.src
        assert 'class="logo"' in self.src

    def test_nav_has_four_links(self):
        """Nav must include 選股方法, 市場概覽, 最新精選, 高危名單."""
        for label in ["選股方法", "市場概覽", "最新精選", "高危名單"]:
            assert label in self.src, f"Nav link '{label}' missing from about.html"

    def test_nav_links_have_i18n_attributes(self):
        """Nav links must carry data-zh-hk / data-zh-cn / data-en for translation."""
        assert 'data-zh-hk=' in self.src
        assert 'data-zh-cn=' in self.src
        assert 'data-en=' in self.src

    def test_lang_buttons_use_setlang(self):
        """Language buttons must call setLang(), not the old switchLang()."""
        assert "onclick=\"setLang('zh-hk')\"" in self.src
        assert "onclick=\"setLang('en')\"" in self.src
        assert "switchLang" not in self.src, (
            "about.html still uses old switchLang() — should be setLang()"
        )

    def test_setlang_saves_to_localstorage(self):
        """setLang() in about.html must persist the language."""
        assert "localStorage.setItem('hidh_lang'" in self.src

    def test_setlang_restores_on_load(self):
        """Page load must read localStorage and restore previous language."""
        assert "localStorage.getItem('hidh_lang')" in self.src

    def test_has_three_content_sections(self):
        """Three language sections must be present with correct IDs."""
        for sec_id in ["sec-zh-hk", "sec-zh-cn", "sec-en"]:
            assert f'id="{sec_id}"' in self.src, f"Missing content section id={sec_id}"

    def test_has_avoid_list_section(self):
        """about.html must explain the High-Risk / 高危名單 list."""
        assert "高危名單" in self.src or "High-Risk List" in self.src

    def test_has_footer(self):
        """about.html must have a dark footer."""
        assert 'class="footer"' in self.src


# ══════════════════════════════════════════════════════════════════════════
# 3. privacy.html — header, nav, persistence
# ══════════════════════════════════════════════════════════════════════════

class TestPrivacyHtml:
    def setup_method(self):
        self.src = _read("privacy.html")

    def test_has_header_with_logo(self):
        assert 'class="header"' in self.src
        assert 'class="logo"' in self.src

    def test_nav_has_four_links(self):
        for label in ["選股方法", "市場概覽", "最新精選", "高危名單"]:
            assert label in self.src, f"Nav link '{label}' missing from privacy.html"

    def test_lang_buttons_use_setlang(self):
        assert "onclick=\"setLang('zh-hk')\"" in self.src
        assert "switchLang" not in self.src

    def test_setlang_saves_to_localstorage(self):
        assert "localStorage.setItem('hidh_lang'" in self.src

    def test_setlang_restores_on_load(self):
        assert "localStorage.getItem('hidh_lang')" in self.src

    def test_has_three_content_sections(self):
        for sec_id in ["sec-zh-hk", "sec-zh-cn", "sec-en"]:
            assert f'id="{sec_id}"' in self.src, f"Missing content section id={sec_id}"

    def test_has_footer(self):
        assert 'class="footer"' in self.src
