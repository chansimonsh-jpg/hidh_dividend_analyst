"""
test_scoring.py — Comprehensive tests for the dividend-stock scoring pipeline.

Covers:
  • All 5 scoring dimensions (dividend quality, valuation, financial health,
    growth, technical) for daily_importer, batch_importer, and screener.
  • Cross-module consistency: screener ↔ daily_importer should be identical
    for all inputs except known documented bugs.
  • Known bugs marked xfail so they appear in the report without blocking CI.
  • generate_html rating logic and daily_importer get_status labels.
"""

import sys
import os
import importlib.util

import pytest
import pandas as pd
import numpy as np

# ── Load production modules ────────────────────────────────────────────────
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _load(alias, rel_path):
    """Import a standalone script as a module without running __main__."""
    path = os.path.join(ROOT, rel_path)
    spec = importlib.util.spec_from_file_location(alias, path)
    mod  = importlib.util.module_from_spec(spec)
    sys.modules[alias] = mod
    spec.loader.exec_module(mod)
    return mod


daily    = _load("_daily",    "daily_importer_global_v6.py")
batch    = _load("_batch",    "batch_importer_global_v6.py")
screener = _load("_screener", "screener_global.py")
# generate_html.py is at repo root in git, under web_site/ in the dev working tree
_html_rel = ("generate_html.py"
             if os.path.exists(os.path.join(ROOT, "generate_html.py"))
             else os.path.join("web_site", "generate_html.py"))
html_mod = _load("_html", _html_rel)

# ── Test-data helpers ──────────────────────────────────────────────────────

def _hist(years_divs: dict, n_closes: int = 30, close: float = 100.0) -> pd.DataFrame:
    """
    Synthetic df_hist with columns Date / Close / Dividend_Amount.
    years_divs: {year: total_dividend_for_that_year}  (one row per year)
    n_closes: additional zero-dividend close rows (need ≥15 for RSI)
    """
    rows = []
    for yr, div in sorted(years_divs.items()):
        rows.append({"Date": pd.Timestamp(f"{yr}-06-15"),
                     "Dividend_Amount": float(div), "Close": close})
    for i in range(n_closes):
        rows.append({"Date": pd.Timestamp(f"2024-{(i % 12) + 1:02d}-{(i % 28) + 1:02d}"),
                     "Dividend_Amount": 0.0, "Close": close + i * 0.1})
    return pd.DataFrame(rows)


def _info(**overrides) -> dict:
    """Base info dict with safe defaults; override any field via kwargs."""
    base = {
        "payoutRatio":            0.50,
        "trailingEps":            2.0,
        "dividendRate":           1.0,
        "freeCashflow":           2e9,
        "sharesOutstanding":      1e9,
        "ebitda":                 5e9,
        "totalDebt":              5e9,
        "totalCash":              1e9,
        "ebit":                   1e9,
        "interestExpense":        0.1e9,
        "currentRatio":           1.5,
        "priceToBook":            1.0,
        "earningsGrowth":         0.05,
        "fiftyTwoWeekLow":        80.0,
        "fiftyTwoWeekHigh":      120.0,
        "currentPrice":           90.0,
    }
    base.update(overrides)
    return base


# ── Yield / PE band params (shared across tests) ───────────────────────────
Y_AVG, Y_STD   = 5.0, 1.0   # y_buy=6.5, y_sell=3.5
PE_AVG, PE_STD = 15.0, 3.0  # pe_buy=10.5, pe_sell=19.5
HIST_STABLE    = _hist({2021: 1.0, 2022: 1.05, 2023: 1.10, 2024: 1.15})
HIST_CUT       = _hist({2021: 1.0, 2022: 0.80, 2023: 0.85, 2024: 0.90})  # 2022 cut
HIST_SHORT     = _hist({2023: 1.0})          # only 1 year → <2 rows → 1pt
HIST_NONE      = None


# ══════════════════════════════════════════════════════════════════════════
# 1.  DIVIDEND QUALITY  (max 30 pts)
# ══════════════════════════════════════════════════════════════════════════

class TestDividendQuality:
    """Tests for score_dividend_quality in daily_importer (reference)."""

    # ── Yield signal (0-10 pts) ──────────────────────────────────────────

    def test_yield_exceptional(self):
        """Yield above buy band → 10 pts."""
        info = _info(payoutRatio=0.50, freeCashflow=2e9, sharesOutstanding=1e9, dividendRate=1.0)
        pts = daily.score_dividend_quality(info, c_yield=7.0, y_avg=Y_AVG, y_std=Y_STD,
                                            df_hist=HIST_STABLE)
        # Yield signal should be 10 (>6.5), payout 10, fcf 8, history 2
        assert pts == 30

    def test_yield_healthy(self):
        """Yield between avg and buy band → 7 pts yield component."""
        info = _info(payoutRatio=0.50, freeCashflow=2e9, sharesOutstanding=1e9, dividendRate=1.0)
        pts = daily.score_dividend_quality(info, c_yield=5.5, y_avg=Y_AVG, y_std=Y_STD,
                                            df_hist=HIST_STABLE)
        # Yield 7, payout 10, fcf 8, history 2 = 27
        assert pts == 27

    def test_yield_adequate(self):
        """Yield between sell band and avg → 4 pts yield component."""
        info = _info(payoutRatio=0.50, freeCashflow=2e9, sharesOutstanding=1e9, dividendRate=1.0)
        pts = daily.score_dividend_quality(info, c_yield=4.5, y_avg=Y_AVG, y_std=Y_STD,
                                            df_hist=HIST_STABLE)
        # Yield 4, payout 10, fcf 8, history 2 = 24
        assert pts == 24

    def test_yield_low(self):
        """Yield below sell band → 1 pt yield component."""
        info = _info(payoutRatio=0.50, freeCashflow=2e9, sharesOutstanding=1e9, dividendRate=1.0)
        pts = daily.score_dividend_quality(info, c_yield=3.0, y_avg=Y_AVG, y_std=Y_STD,
                                            df_hist=HIST_STABLE)
        # Yield 1, payout 10, fcf 8, history 2 = 21
        assert pts == 21

    # ── Payout ratio (0-10 pts) ──────────────────────────────────────────

    @pytest.mark.parametrize("pr,expected_payout_pts", [
        (0.30, 10),   # conservative payout
        (0.50, 10),   # exactly at boundary
        (0.51, 8),    # just over 50%
        (0.65, 8),    # exactly at 65% boundary
        (0.66, 5),    # 65-80 range
        (0.80, 5),    # exactly at 80% boundary
        (0.81, 2),    # 80-95 range
        (0.95, 2),    # exactly at 95% boundary
        (0.96, 0),    # 95-100% range — daily gives 0 (explicit)
        (1.00, 0),    # exactly 100%
        (1.01, 0),    # >100%
        (1.50, 0),    # well over 100%
    ])
    def test_payout_ratio_daily(self, pr, expected_payout_pts):
        """Payout ratio scoring in daily_importer."""
        # Use fixed yield (at buy band) + fcf + history so we can isolate payout pts
        # yield → 10, fcf → 8, history → 2 ; only payout varies
        info = _info(payoutRatio=pr, freeCashflow=2e9, sharesOutstanding=1e9, dividendRate=1.0)
        pts = daily.score_dividend_quality(info, c_yield=7.0, y_avg=Y_AVG, y_std=Y_STD,
                                            df_hist=HIST_STABLE)
        # yield=10, payout=expected_payout_pts, fcf=8, history=2
        assert pts == 10 + expected_payout_pts + 8 + 2

    def test_payout_missing_daily(self):
        """Missing payout ratio (pr < 0, can't compute from EPS either) → 2 pts in daily."""
        # trailingEps=0 prevents implied payout from being computed (div/eps with eps=0 → -1)
        info = _info(payoutRatio=None, trailingEps=0,
                     freeCashflow=2e9, sharesOutstanding=1e9, dividendRate=1.0)
        pts = daily.score_dividend_quality(info, c_yield=7.0, y_avg=Y_AVG, y_std=Y_STD,
                                            df_hist=HIST_STABLE)
        # yield=10, payout_missing=2, fcf=8 (2e9/(1e9*1.0)=2x), history=2 = 22
        assert pts == 22

    # ── FCF coverage (0-8 pts) ───────────────────────────────────────────

    @pytest.mark.parametrize("fcf_cov,expected_fcf_pts", [
        (2.0,  8),   # exactly 2x
        (3.0,  8),   # above 2x
        (1.5,  6),   # exactly 1.5x
        (1.8,  6),   # 1.5-2.0 range
        (1.0,  4),   # exactly 1.0x
        (1.2,  4),   # 1.0-1.5 range
        (0.8,  1),   # below 1.0x
    ])
    def test_fcf_coverage_daily(self, fcf_cov, expected_fcf_pts):
        """FCF coverage scoring in daily_importer."""
        div_paid = 1e9   # shares=1e9 * dividendRate=1.0
        fcf = fcf_cov * div_paid
        info = _info(payoutRatio=0.50, freeCashflow=fcf, sharesOutstanding=1e9, dividendRate=1.0)
        pts = daily.score_dividend_quality(info, c_yield=7.0, y_avg=Y_AVG, y_std=Y_STD,
                                            df_hist=HIST_STABLE)
        # yield=10, payout=10, fcf=expected, history=2
        assert pts == 10 + 10 + expected_fcf_pts + 2

    def test_negative_fcf_daily_penalizes(self):
        """Negative FCF → explicit 0 pts penalty in daily_importer."""
        info = _info(payoutRatio=0.50, freeCashflow=-1e9, sharesOutstanding=1e9, dividendRate=1.0)
        pts = daily.score_dividend_quality(info, c_yield=7.0, y_avg=Y_AVG, y_std=Y_STD,
                                            df_hist=HIST_STABLE)
        # yield=10, payout=10, fcf=0 (penalty), history=2 = 22
        assert pts == 22

    def test_missing_fcf_daily(self):
        """FCF=0 with no dividend rate data → 2 pts (missing data) in daily."""
        info = _info(payoutRatio=0.50, freeCashflow=0, sharesOutstanding=0, dividendRate=0)
        pts = daily.score_dividend_quality(info, c_yield=7.0, y_avg=Y_AVG, y_std=Y_STD,
                                            df_hist=HIST_STABLE)
        # yield=10, payout=10, fcf_missing=2, history=2 = 24
        assert pts == 24

    # ── Dividend history / cut detection (0-2 pts) ───────────────────────

    def test_no_cut_history(self):
        """Stable or growing dividends → 2 pts."""
        info = _info(payoutRatio=0.50, freeCashflow=2e9, sharesOutstanding=1e9, dividendRate=1.0)
        pts = daily.score_dividend_quality(info, c_yield=7.0, y_avg=Y_AVG, y_std=Y_STD,
                                            df_hist=HIST_STABLE)
        assert pts == 30  # perfect score

    def test_cut_detected(self):
        """Any year < 90% of prior → 0 pts for history component."""
        info = _info(payoutRatio=0.50, freeCashflow=2e9, sharesOutstanding=1e9, dividendRate=1.0)
        pts = daily.score_dividend_quality(info, c_yield=7.0, y_avg=Y_AVG, y_std=Y_STD,
                                            df_hist=HIST_CUT)
        # yield=10, payout=10, fcf=8, history=0 (cut detected) = 28
        assert pts == 28

    def test_no_history(self):
        """No history data → 1 pt."""
        info = _info(payoutRatio=0.50, freeCashflow=2e9, sharesOutstanding=1e9, dividendRate=1.0)
        pts = daily.score_dividend_quality(info, c_yield=7.0, y_avg=Y_AVG, y_std=Y_STD,
                                            df_hist=HIST_NONE)
        # yield=10, payout=10, fcf=8, history=1 (no data) = 29
        assert pts == 29

    def test_short_history_one_year(self):
        """Only 1 year of dividends → 1 pt (< 2 rows)."""
        info = _info(payoutRatio=0.50, freeCashflow=2e9, sharesOutstanding=1e9, dividendRate=1.0)
        pts = daily.score_dividend_quality(info, c_yield=7.0, y_avg=Y_AVG, y_std=Y_STD,
                                            df_hist=HIST_SHORT)
        # yield=10, payout=10, fcf=8, history=1 = 29
        assert pts == 29


# ══════════════════════════════════════════════════════════════════════════
# 2.  VALUATION  (max 25 pts)
# ══════════════════════════════════════════════════════════════════════════

class TestValuation:
    """Tests for score_valuation in daily_importer (reference)."""

    # ── PE position (0-8 pts) ────────────────────────────────────────────

    @pytest.mark.parametrize("c_pe,expected_pe_pts", [
        (10.0, 8),   # ≤ pe_buy=10.5 → cheap
        (10.5, 8),   # exactly at pe_buy
        (12.0, 6),   # pe_buy < pe ≤ pe_avg=15
        (15.0, 6),   # exactly at pe_avg
        (17.0, 3),   # pe_avg < pe ≤ pe_sell=19.5
        (19.5, 3),   # exactly at pe_sell
        (20.0, 0),   # above pe_sell → expensive
    ])
    def test_pe_scoring_daily(self, c_pe, expected_pe_pts):
        """PE position scoring in daily_importer."""
        # pb=0.9 → 5pts, spread = 7.0 - 4.3 = 2.7 → 8pts
        info = _info(priceToBook=0.9)
        pts = daily.score_valuation(info, c_yield=7.0, y_avg=Y_AVG, y_std=Y_STD,
                                     c_pe=c_pe, pe_avg=PE_AVG, pe_std=PE_STD, rfr=4.3)
        # PE=expected, PB=5, spread=8
        assert pts == expected_pe_pts + 5 + 8

    # ── P/B ratio (0-7 pts) ──────────────────────────────────────────────

    @pytest.mark.parametrize("pb,expected_pb_pts", [
        (0.5,  7),    # deep value ≤ 0.8
        (0.8,  7),    # exactly at 0.8
        (0.9,  5),    # 0.8 < pb ≤ 1.2
        (1.2,  5),    # exactly at 1.2
        (1.5,  3),    # 1.2 < pb ≤ 2.0
        (2.0,  3),    # exactly at 2.0
        (2.5,  1),    # 2.0 < pb ≤ 3.5
        (3.5,  1),    # exactly at 3.5
        (4.0,  0),    # above 3.5
        (-1.0, 2),    # negative → missing data
    ])
    def test_pb_scoring_daily(self, pb, expected_pb_pts):
        """P/B scoring in daily_importer (and screener — they share thresholds)."""
        info = _info(priceToBook=pb)
        pts = daily.score_valuation(info, c_yield=7.0, y_avg=Y_AVG, y_std=Y_STD,
                                     c_pe=PE_AVG, pe_avg=PE_AVG, pe_std=PE_STD, rfr=4.3)
        # PE at avg → 6pts, PB=expected, spread(7.0-4.3=2.7)→8pts
        assert pts == 6 + expected_pb_pts + 8

    # ── Yield spread (0-10 pts) ──────────────────────────────────────────

    @pytest.mark.parametrize("c_yield,rfr,expected_spread_pts", [
        (7.3, 4.3, 10),   # spread=3.0 exactly → 10
        (8.0, 4.3, 10),   # spread>3.0 → 10
        (6.3, 4.3, 8),    # spread=2.0 → 8
        (5.3, 4.3, 6),    # spread=1.0 → 6
        (4.3, 4.3, 4),    # spread=0.0 → 4
        (3.3, 4.3, 2),    # spread=-1.0 exactly → 2 (≥-1.0)
        (2.3, 4.3, 0),    # spread=-2.0, below -1.0 → 0
    ])
    def test_yield_spread_daily(self, c_yield, rfr, expected_spread_pts):
        """Yield spread scoring in daily_importer."""
        info = _info(priceToBook=1.0)  # pb=1.0 → 5pts
        pts = daily.score_valuation(info, c_yield=c_yield, y_avg=Y_AVG, y_std=Y_STD,
                                     c_pe=PE_AVG, pe_avg=PE_AVG, pe_std=PE_STD, rfr=rfr)
        # PE at avg → 6, PB=5, spread=expected
        assert pts == 6 + 5 + expected_spread_pts

    def test_yield_spread_below_minus1(self):
        """Spread < -1.0 → 0 pts."""
        info = _info(priceToBook=1.0)
        pts = daily.score_valuation(info, c_yield=2.0, y_avg=Y_AVG, y_std=Y_STD,
                                     c_pe=PE_AVG, pe_avg=PE_AVG, pe_std=PE_STD, rfr=4.3)
        # spread = 2.0 - 4.3 = -2.3 → 0 pts
        assert pts == 6 + 5 + 0


# ══════════════════════════════════════════════════════════════════════════
# 3.  FINANCIAL HEALTH  (max 25 pts)
# ══════════════════════════════════════════════════════════════════════════

class TestFinancialHealth:
    """Tests for score_financial_health in daily_importer (reference)."""

    # ── Net Debt / EBITDA (0-10 pts) ────────────────────────────────────

    @pytest.mark.parametrize("total_debt,total_cash,ebitda,expected_nd_pts", [
        (1e9, 0.5e9, 1e9,  10),  # nd/ebitda=0.5 ≤ 1.0
        (2e9, 0,     1e9,  8),   # nd/ebitda=2.0 ≤ 2.0
        (3e9, 0,     1e9,  6),   # nd/ebitda=3.0 ≤ 3.0
        (4.5e9, 0,   1e9,  3),   # nd/ebitda=4.5 ≤ 4.5
        (5e9, 0,     1e9,  0),   # nd/ebitda=5.0 > 4.5
        (0,   1e9,   0,    10),  # net cash (ebitda=0, nd<0) → top score
        (1e9, 0,     0,    2),   # has debt, no ebitda → low score
    ])
    def test_nd_ebitda_daily(self, total_debt, total_cash, ebitda, expected_nd_pts):
        """Net Debt/EBITDA scoring in daily_importer."""
        # interest coverage: ebit=1e9, ie=0.1e9 → ic=10 → 9pts
        # current ratio: 1.5 → 5pts
        info = _info(totalDebt=total_debt, totalCash=total_cash, ebitda=ebitda,
                     ebit=1e9, interestExpense=0.1e9, currentRatio=1.5)
        pts = daily.score_financial_health(info)
        assert pts == expected_nd_pts + 9 + 5

    # ── Interest coverage (0-9 pts) ──────────────────────────────────────

    @pytest.mark.parametrize("ebit,ie,expected_ic_pts", [
        (8e9,  1e9, 9),   # ic=8 → 9pts
        (8e9,  1e9, 9),
        (5e9,  1e9, 7),   # ic=5 → 7pts
        (3e9,  1e9, 4),   # ic=3 → 4pts
        (1.5e9,1e9, 2),   # ic=1.5 → 2pts
        (1e9,  1e9, 0),   # ic=1.0 < 1.5 → 0pts
        (-1e9, 1e9, 0),   # negative EBIT → 0pts
    ])
    def test_interest_coverage_daily(self, ebit, ie, expected_ic_pts):
        """Interest coverage scoring in daily_importer."""
        # nd/ebitda: set ebitda high so nd is small → 10pts
        # current ratio: 1.5 → 5pts
        info = _info(totalDebt=1e9, totalCash=0.5e9, ebitda=10e9,
                     ebit=ebit, interestExpense=ie, currentRatio=1.5)
        pts = daily.score_financial_health(info)
        assert pts == 10 + expected_ic_pts + 5

    def test_missing_interest_coverage_daily(self):
        """ebit>0 but interestExpense=0 → 2 pts (missing data) in daily."""
        info = _info(totalDebt=1e9, totalCash=0.5e9, ebitda=10e9,
                     ebit=1e9, interestExpense=0, currentRatio=1.5)
        pts = daily.score_financial_health(info)
        assert pts == 10 + 2 + 5

    # ── Current ratio (0-6 pts) ──────────────────────────────────────────

    @pytest.mark.parametrize("cr,expected_cr_pts", [
        (2.0,  6),
        (2.5,  6),
        (1.5,  5),
        (1.9,  5),
        (1.0,  3),
        (1.4,  3),
        (0.8,  0),   # below 1.0 → 0
        (-1.0, 1),   # missing (cr < 0) → 1 pt
    ])
    def test_current_ratio_daily(self, cr, expected_cr_pts):
        """Current ratio scoring in daily_importer."""
        # nd/ebitda → 10, interest coverage → 9
        info = _info(totalDebt=1e9, totalCash=0.5e9, ebitda=10e9,
                     ebit=8e9, interestExpense=1e9, currentRatio=cr)
        pts = daily.score_financial_health(info)
        assert pts == 10 + 9 + expected_cr_pts


# ══════════════════════════════════════════════════════════════════════════
# 4.  GROWTH  (max 10 pts)
# ══════════════════════════════════════════════════════════════════════════

class TestGrowth:
    """Tests for score_growth in daily_importer (reference)."""

    # ── Dividend growth rate (0-6 pts) ──────────────────────────────────

    @pytest.mark.parametrize("years_divs,expected_dgr_pts", [
        ({2021: 1.0, 2022: 1.05, 2023: 1.11, 2024: 1.26}, 6),   # CAGR=(1.26)^(1/3)-1≈8.0% → ≥8
        ({2021: 1.0, 2022: 1.03, 2023: 1.06, 2024: 1.16}, 5),   # CAGR=(1.16)^(1/3)-1≈5.0% → ≥5
        ({2021: 1.0, 2022: 1.03, 2023: 1.06, 2024: 1.10}, 4),   # CAGR=(1.10)^(1/3)-1≈3.2% → ≥2
        ({2021: 1.0, 2022: 1.00, 2023: 1.00, 2024: 1.00}, 2),   # CAGR=0%
        ({2021: 1.0, 2022: 0.98, 2023: 0.96, 2024: 0.90}, 0),   # CAGR<0
    ])
    def test_dgr_4yr_cagr_daily(self, years_divs, expected_dgr_pts):
        """4-year dividend growth rate (3yr CAGR) in daily_importer."""
        hist = _hist(years_divs)
        info = _info(earningsGrowth=0.05)    # earnings: ≥3% → 3pts
        pts = daily.score_growth(info, hist)
        assert pts == expected_dgr_pts + 3

    def test_dgr_2yr_cagr_daily(self):
        """Only 2 years of history → simple 1yr CAGR."""
        hist = _hist({2023: 1.0, 2024: 1.10})   # 10% growth
        info = _info(earningsGrowth=0.05)
        pts = daily.score_growth(info, hist)
        # dgr=10% → 6pts; earnings≥3% → 3pts
        assert pts == 6 + 3

    # ── Earnings growth (0-4 pts) ────────────────────────────────────────

    @pytest.mark.parametrize("eg,expected_eg_pts", [
        (0.10,  4),
        (0.15,  4),
        (0.03,  3),
        (0.09,  3),
        (0.02,  2),
        (-0.05, 0),
        (None,  1),   # missing → 1 pt in daily/screener
    ])
    def test_earnings_growth_daily(self, eg, expected_eg_pts):
        """Earnings growth scoring in daily_importer."""
        # Use stable dividends for 0% DGR → 2 pts DGR component
        hist = _hist({2021: 1.0, 2022: 1.0, 2023: 1.0, 2024: 1.0})
        info = _info(earningsGrowth=eg)
        pts = daily.score_growth(info, hist)
        assert pts == 2 + expected_eg_pts

    def test_earnings_growth_zero_not_treated_as_missing(self):
        """earningsGrowth=0.0 (flat earnings) correctly scores 2pts (≥0), not missing-data 1pt."""
        hist = _hist({2021: 1.0, 2022: 1.0, 2023: 1.0, 2024: 1.0})
        info = _info(earningsGrowth=0.0)
        pts = daily.score_growth(info, hist)
        assert pts == 2 + 2  # dgr=0%→2, eg=0%→2

    def test_no_history_growth_daily(self):
        """No history → dgr=0 → 2 pts; earnings missing → 1 pt = 3 pts total."""
        info = _info(earningsGrowth=None)
        pts = daily.score_growth(info, df_hist=None)
        assert pts == 2 + 1


# ══════════════════════════════════════════════════════════════════════════
# 5.  TECHNICAL  (max 10 pts)
# ══════════════════════════════════════════════════════════════════════════

class TestTechnical:
    """Tests for score_technical in daily_importer (reference)."""

    # ── 52-week position (0-6 pts) ───────────────────────────────────────

    @pytest.mark.parametrize("price,low,high,expected_pos_pts", [
        (82.0, 80.0, 120.0, 6),   # pos=0.05 ≤ 0.25 → 6
        (90.0, 80.0, 120.0, 6),   # pos=0.25 exactly → ≤0.25 → 6
        (95.0, 80.0, 120.0, 5),   # pos=0.375 > 0.25, ≤ 0.40 → 5
        (100.0,80.0, 120.0, 3),   # pos=0.50 ≤ 0.60 → 3
        (112.0,80.0, 120.0, 1),   # pos=0.80 exactly ≤ 0.80 → 1
        (115.0,80.0, 120.0, 0),   # pos=0.875 > 0.80 → 0
        (0,    0,    0,     1),   # missing data → 1 pt
    ])
    def test_52w_position_daily(self, price, low, high, expected_pos_pts):
        """52-week position scoring in daily_importer."""
        info = _info(fiftyTwoWeekLow=low, fiftyTwoWeekHigh=high, currentPrice=price)
        # Use short history (< 15 rows) so RSI gives 1 pt
        short_hist = _hist({}, n_closes=5)
        pts = daily.score_technical(short_hist, info)
        assert pts == expected_pos_pts + 1

    def test_52w_position_0_25_boundary(self):
        """pos=0.25 exactly is ≤0.25 → 6 pts."""
        info = _info(fiftyTwoWeekLow=80.0, fiftyTwoWeekHigh=120.0, currentPrice=90.0)
        short_hist = _hist({}, n_closes=5)
        pts = daily.score_technical(short_hist, info)
        assert pts == 6 + 1

    # ── RSI (0-4 pts) ────────────────────────────────────────────────────

    def _hist_with_declining_prices(self, n=30, start=100.0, step=-1.0):
        """Generate n rows of steadily declining prices → RSI will be low (oversold)."""
        rows = [{"Date": pd.Timestamp(f"2024-01-{i+1:02d}") if i < 31 else
                          pd.Timestamp(f"2024-02-{i-30:02d}"),
                 "Close": max(0.1, start + i * step),
                 "Dividend_Amount": 0.0}
                for i in range(n)]
        return pd.DataFrame(rows)

    def _hist_with_rising_prices(self, n=30, start=100.0, step=1.0):
        rows = [{"Date": pd.Timestamp(f"2024-01-{i+1:02d}") if i < 31 else
                          pd.Timestamp(f"2024-02-{i-30:02d}"),
                 "Close": start + i * step,
                 "Dividend_Amount": 0.0}
                for i in range(n)]
        return pd.DataFrame(rows)

    def test_rsi_oversold(self):
        """Steadily declining prices → RSI ≤ 30 → 4 pts RSI component."""
        info = _info(fiftyTwoWeekLow=70.0, fiftyTwoWeekHigh=120.0, currentPrice=75.0)
        hist = self._hist_with_declining_prices(n=30, start=100.0, step=-2.0)
        pts = daily.score_technical(hist, info)
        # pos = (75-70)/(120-70) = 0.10 ≤ 0.25 → 6pts
        # RSI very low from declining → 4pts
        assert pts == 6 + 4

    def test_rsi_missing_not_enough_rows(self):
        """Fewer than 15 close rows → RSI not computed → 1 pt fallback."""
        info = _info(fiftyTwoWeekLow=80.0, fiftyTwoWeekHigh=120.0, currentPrice=90.0)
        short = _hist({}, n_closes=10)  # <15 rows
        pts = daily.score_technical(short, info)
        # pos ≤ 0.25 → 6, rsi_missing → 1
        assert pts == 6 + 1


# ══════════════════════════════════════════════════════════════════════════
# 6.  TOTAL SCORE INTEGRITY
# ══════════════════════════════════════════════════════════════════════════

class TestComputeScore:
    """compute_score sums all 5 dimensions correctly."""

    def test_total_score_is_sum_of_dimensions(self):
        """Score_總分_100 must equal sum of 5 sub-scores."""
        info = _info()
        result = daily.compute_score(
            info, c_yield=7.0, y_avg=Y_AVG, y_std=Y_STD,
            c_pe=PE_AVG, pe_avg=PE_AVG, pe_std=PE_STD,
            df_hist=HIST_STABLE, rfr=4.3,
        )
        expected_total = round(
            result["Score_股息質量_30"] +
            result["Score_估值_25"] +
            result["Score_財務健康_25"] +
            result["Score_增長潛力_10"] +
            result["Score_技術面_10"],
            1,
        )
        assert result["Score_總分_100"] == expected_total

    def test_max_possible_score(self):
        """A perfect stock should score exactly 100 pts."""
        # Dividend rows first (earlier dates) so they don't affect RSI tail
        hist_rows = []
        for yr, div in sorted({2021: 1.0, 2022: 1.09, 2023: 1.19, 2024: 1.30}.items()):
            hist_rows.append({"Date": pd.Timestamp(f"{yr}-06-15"),
                              "Close": 90.0, "Dividend_Amount": div})
        # Declining prices at the end → RSI tail is oversold (≤30) → 4pts
        for i in range(35):
            hist_rows.append({
                "Date": pd.Timestamp("2025-01-01") + pd.Timedelta(days=i),
                "Close": max(1.0, 90.0 - i * 2.5),   # 90 → ~5, pure decline
                "Dividend_Amount": 0.0,
            })
        hist = pd.DataFrame(hist_rows)

        info = _info(
            payoutRatio=0.40, freeCashflow=3e9,   # fcf_cov=3x → 8pts
            sharesOutstanding=1e9, dividendRate=1.0,
            priceToBook=0.5,                       # ≤0.8 → 7pts
            ebitda=10e9, totalDebt=1e9, totalCash=0.5e9,  # nd/ebitda=0.05 → 10pts
            ebit=10e9, interestExpense=0.5e9,       # ic=20 → 9pts
            currentRatio=2.5,                       # ≥2.0 → 6pts
            earningsGrowth=0.15,                    # ≥0.10 → 4pts
            fiftyTwoWeekLow=12.5, fiftyTwoWeekHigh=100.0,
            currentPrice=15.0,   # pos=(15-12.5)/(100-12.5)=0.029 ≤ 0.25 → 6pts
        )
        # c_yield=8.0: spread=8.0-4.3=3.7 ≥3.0 → 10pts; yield>6.5 → 10pts
        # c_pe=10.0 ≤ pe_buy=10.5 → 8pts
        result = daily.compute_score(
            info, c_yield=8.0, y_avg=Y_AVG, y_std=Y_STD,
            c_pe=10.0, pe_avg=PE_AVG, pe_std=PE_STD,
            df_hist=hist, rfr=4.3,
        )
        assert result["Score_總分_100"] == 100

    def test_minimum_score_non_negative(self):
        """A terrible stock should score ≥ 0 (no negative scores)."""
        info = _info(
            payoutRatio=2.0,       # >1 → 0pts
            freeCashflow=-1e9,     # negative → 0pts
            sharesOutstanding=1e9,
            priceToBook=5.0,       # > 3.5 → 0pts
            ebitda=1e9, totalDebt=10e9, totalCash=0,  # nd/ebitda=10 → 0pts
            ebit=-1e9,             # negative → 0pts
            interestExpense=1e9,
            currentRatio=0.5,     # < 1 → 0pts
            earningsGrowth=-0.20,  # negative → 0pts
            fiftyTwoWeekLow=80.0, fiftyTwoWeekHigh=120.0, currentPrice=119.0,  # near top
        )
        hist = _hist({2021: 1.0, 2022: 0.7, 2023: 0.6, 2024: 0.5})  # cut
        result = daily.compute_score(
            info, c_yield=2.0, y_avg=Y_AVG, y_std=Y_STD,
            c_pe=30.0, pe_avg=PE_AVG, pe_std=PE_STD,
            df_hist=hist, rfr=4.3,
        )
        assert result["Score_總分_100"] >= 0


# ══════════════════════════════════════════════════════════════════════════
# 7.  RATING LABELS
# ══════════════════════════════════════════════════════════════════════════

class TestRatingLabels:
    """Tests for get_status (daily_importer) and get_rating_key (generate_html)."""

    @pytest.mark.parametrize("score,expected_label", [
        (75,  "🟢🟢 強力買入"),
        (80,  "🟢🟢 強力買入"),
        (60,  "🟢 值得關注"),
        (74,  "🟢 值得關注"),
        (45,  "⚖️ 觀望"),
        (59,  "⚖️ 觀望"),
        (30,  "🟡 偏弱"),
        (44,  "🟡 偏弱"),
        (29,  "🔴 避開"),
        (0,   "🔴 避開"),
    ])
    def test_get_status_daily(self, score, expected_label):
        """daily_importer.get_status maps scores to 5-tier labels."""
        assert daily.get_status(score) == expected_label

    @pytest.mark.parametrize("score,expected_key", [
        (75, "strong"),
        (80, "strong"),
        (60, "watch"),   # new Watch threshold
        (74, "watch"),
        (59, "hold"),    # 50-59 is now "hold", not "watch"
        (50, "hold"),
        (49, "hold"),
        (0,  "hold"),
    ])
    def test_get_rating_key_html(self, score, expected_key):
        """generate_html.get_rating_key maps scores to 3-tier keys (Watch threshold = 60)."""
        assert html_mod.get_rating_key(score) == expected_key


# ══════════════════════════════════════════════════════════════════════════
# 8.  CROSS-MODULE CONSISTENCY  (screener ↔ daily)
# ══════════════════════════════════════════════════════════════════════════
# The screener and daily_importer are supposed to use identical scoring.
# Any discrepancy is a bug.  The tests below confirm agreement on normal
# inputs and expose known divergences as xfail.

class TestCrossModuleConsistency:
    """screener vs daily_importer scoring should be identical for all inputs."""

    def _assert_same(self, fn_daily, fn_screener, *args, **kwargs):
        d = fn_daily(*args, **kwargs)
        s = fn_screener(*args, **kwargs)
        assert d == s, f"daily={d}, screener={s}"

    def test_div_quality_stable_match(self):
        """Nominal case: screener and daily agree."""
        info = _info(payoutRatio=0.50, freeCashflow=2e9, sharesOutstanding=1e9, dividendRate=1.0)
        self._assert_same(
            daily.score_dividend_quality, screener.score_dividend_quality,
            info, 7.0, Y_AVG, Y_STD, HIST_STABLE,
        )

    def test_valuation_match(self):
        info = _info(priceToBook=0.7)
        self._assert_same(
            daily.score_valuation, screener.score_valuation,
            info, 7.0, Y_AVG, Y_STD, 15.0, PE_AVG, PE_STD, 4.3,
        )

    def test_financial_health_match(self):
        info = _info(totalDebt=1e9, totalCash=0.5e9, ebitda=10e9,
                     ebit=5e9, interestExpense=1e9, currentRatio=1.5)
        self._assert_same(
            daily.score_financial_health, screener.score_financial_health, info,
        )

    def test_growth_match(self):
        info = _info(earningsGrowth=0.08)
        self._assert_same(
            daily.score_growth, screener.score_growth, info, HIST_STABLE,
        )

    def test_technical_match(self):
        info = _info(fiftyTwoWeekLow=80.0, fiftyTwoWeekHigh=120.0, currentPrice=90.0)
        self._assert_same(
            daily.score_technical, screener.score_technical, HIST_STABLE, info,
        )

    def test_payout_high_near_100pct_screener_matches_daily(self):
        """pr=0.98 → both daily and screener give 0 payout pts after fix."""
        info = _info(payoutRatio=0.98, freeCashflow=2e9, sharesOutstanding=1e9, dividendRate=1.0)
        d = daily.score_dividend_quality(info, 7.0, Y_AVG, Y_STD, HIST_STABLE)
        s = screener.score_dividend_quality(info, 7.0, Y_AVG, Y_STD, HIST_STABLE)
        assert d == s


# ══════════════════════════════════════════════════════════════════════════
# 9.  BATCH IMPORTER BUGS
# ══════════════════════════════════════════════════════════════════════════
# batch_importer is aligned with screener/daily for logic bugs.
# Missing-data defaults remain intentionally more generous in batch
# (designed for initial import when data may be incomplete).

class TestBatchAlignment:
    """batch_importer scoring is now aligned with screener/daily on logic bugs."""

    def test_negative_fcf_batch_penalizes(self):
        """Negative FCF → batch now gives 0pts (same as daily/screener)."""
        info = _info(payoutRatio=0.50, freeCashflow=-1e9, sharesOutstanding=1e9, dividendRate=1.0)
        b = batch.score_dividend_quality(info, 7.0, Y_AVG, Y_STD, HIST_STABLE)
        d = daily.score_dividend_quality(info, 7.0, Y_AVG, Y_STD, HIST_STABLE)
        assert b == d

    def test_pb_threshold_batch_aligned(self):
        """pb=0.9 → batch now gives 5pts (same as daily/screener, threshold ≤0.8 for 7pts)."""
        info = _info(priceToBook=0.9)
        b = batch.score_valuation(info, 7.0, Y_AVG, Y_STD, PE_AVG, PE_AVG, PE_STD, 4.3)
        d = daily.score_valuation(info, 7.0, Y_AVG, Y_STD, PE_AVG, PE_AVG, PE_STD, 4.3)
        assert b == d

    def test_negative_ebit_batch_penalizes(self):
        """Negative EBIT → batch now gives 0pts (same as daily/screener)."""
        info = _info(totalDebt=1e9, totalCash=0.5e9, ebitda=10e9,
                     ebit=-1e9, interestExpense=1e9, currentRatio=1.5)
        b = batch.score_financial_health(info)
        d = daily.score_financial_health(info)
        assert b == d

    def test_batch_missing_data_more_lenient_than_daily(self):
        """batch systematically assigns higher default scores for missing data.

        This is intentional but means batch-initialized scores are not
        directly comparable to daily-recalculated ones when data is absent.
        """
        # All optionals set to None/missing
        sparse = {
            "payoutRatio": None, "trailingEps": 0, "dividendRate": 0,
            "freeCashflow": 0, "sharesOutstanding": 0,
            "ebitda": 0, "totalDebt": 0, "totalCash": 0,
            "ebit": 0, "interestExpense": 0,
            "currentRatio": None,
            "priceToBook": None,
            "earningsGrowth": None,
            "fiftyTwoWeekLow": 0, "fiftyTwoWeekHigh": 0, "currentPrice": 0,
        }
        b = batch.score_financial_health(sparse)
        d = daily.score_financial_health(sparse)
        # Both non-negative; batch should be >= daily (more lenient)
        assert b >= d
        assert b > 0 and d >= 0


# ══════════════════════════════════════════════════════════════════════════
# 10. GENERATE_HTML RATING vs DAILY_IMPORTER LABEL DIVERGENCE
# ══════════════════════════════════════════════════════════════════════════

class TestRatingConsistency:
    """HTML get_rating_key and daily_importer get_status now use the same thresholds."""

    @pytest.mark.parametrize("score", [50, 55, 59])
    def test_score_50_to_59_is_hold_in_both(self, score):
        """Scores 50-59 are consistently 'hold/觀望' in both HTML and Excel."""
        assert html_mod.get_rating_key(score) == "hold"
        assert "值得關注" not in daily.get_status(score)

    def test_score_60_plus_is_watch_in_both(self):
        """Scores 60-74 are consistently 'watch/值得關注' in both HTML and Excel."""
        for score in (60, 65, 74):
            assert html_mod.get_rating_key(score) == "watch"
            assert "值得關注" in daily.get_status(score)

    def test_score_75_plus_is_strong_in_both(self):
        """Scores ≥75 are consistently 'strong/強力買入' in both HTML and Excel."""
        for score in (75, 80, 95):
            assert html_mod.get_rating_key(score) == "strong"
            assert "強力買入" in daily.get_status(score)
