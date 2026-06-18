#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
export_tickers_json.py
======================
一次性工具：將 batch_importer 的硬編碼 TICKERS_CONFIG 匯出成
tickers_config.json，之後 batch 同 screener 都讀這個 JSON。

使用方法：
    python export_tickers_json.py

只需執行一次。之後直接改 tickers_config.json 即可。
"""

import json, datetime, os

TICKERS_CONFIG = {
    # ── UK ────────────────────────────────────────────────
    "LGEN.L":  ("Legal & General",        "UK"),
    "ABDN.L":  ("abrdn plc",              "UK"),
    "MNG.L":   ("M&G Investments",        "UK"),
    "IMB.L":   ("Imperial Brands",        "UK"),
    "BATS.L":  ("BAT Tobacco",            "UK"),
    "UKW.L":   ("Greencoat UK Wind",      "UK"),
    "ORIT.L":  ("Octopus Renewables",     "UK"),
    "TRIG.L":  ("TRIG Renewables",        "UK"),
    "BSIF.L":  ("Bluefield Solar",        "UK"),
    "LAND.L":  ("Land Securities",        "UK"),
    "AV.L":    ("Aviva plc",              "UK"),
    "HSBA.L":  ("HSBC Holdings",          "UK"),
    "BARC.L":  ("Barclays plc",           "UK"),
    "NWG.L":   ("NatWest Group",          "UK"),
    "LLOY.L":  ("Lloyds Banking Group",   "UK"),
    "SDR.L":   ("Schroders plc",          "UK"),
    "INF.L":   ("Informa plc",            "UK"),
    "OSB.L":   ("OSB Group",              "UK"),
    "BLND.L":  ("British Land",           "UK"),
    "BBOX.L":  ("Tritax Big Box",         "UK"),
    "HMSO.L":  ("Hammerson plc",          "UK"),
    "PHP.L":   ("Primary Health Prop",    "UK"),
    "BYG.L":   ("Big Yellow Group",       "UK"),
    "UU.L":    ("United Utilities",       "UK"),
    "SVT.L":   ("Severn Trent",           "UK"),
    "NG.L":    ("National Grid",          "UK"),
    "SSE.L":   ("SSE plc",               "UK"),
    "HFEL.L":  ("Henderson Far East",     "UK"),
    "MYI.L":   ("Murray International",   "UK"),
    "MRCH.L":  ("Merchants Trust",        "UK"),
    "CTY.L":   ("City of London IT",      "UK"),
    "AAIF.L":  ("Aberdeen Asian Inc",     "UK"),
    "BP.L":    ("BP plc",                "UK"),
    "SHEL.L":  ("Shell plc",             "UK"),
    "AAL.L":   ("Anglo American",         "UK"),
    "RIO.L":   ("Rio Tinto",             "UK"),
    "GSK.L":   ("GSK plc",              "UK"),
    "AZN.L":   ("AstraZeneca",           "UK"),
    "ULVR.L":  ("Unilever plc",          "UK"),
    "BWY.L":   ("Bellway plc",           "UK"),
    "TW.L":    ("Taylor Wimpey",         "UK"),
    "PSN.L":   ("Persimmon plc",         "UK"),
    "VOD.L":   ("Vodafone Group",        "UK"),
    "BT-A.L":  ("BT Group",             "UK"),
    "BA.L":    ("BAE Systems",           "UK"),
    "PSON.L":  ("Pearson plc",           "UK"),
    "WPP.L":   ("WPP plc",              "UK"),
    "OCDO.L":  ("Ocado Group",           "UK"),
    "SMIN.L":  ("Smiths Group plc",      "UK"),
    # ── HK ────────────────────────────────────────────────
    "0005.HK": ("HSBC Asia",             "HK"),
    "2388.HK": ("BOC Hong Kong",         "HK"),
    "1398.HK": ("ICBC",                  "HK"),
    "0939.HK": ("CCB",                   "HK"),
    "3988.HK": ("Bank of China",         "HK"),
    "1288.HK": ("Agricultural Bank",     "HK"),
    "0023.HK": ("Bank of East Asia",     "HK"),
    "0823.HK": ("Link REIT",             "HK"),
    "0016.HK": ("Sun Hung Kai Prop",     "HK"),
    "0012.HK": ("Henderson Land",        "HK"),
    "0101.HK": ("Hang Lung Prop",        "HK"),
    "0778.HK": ("Fortune REIT",          "HK"),
    "2778.HK": ("Champion REIT",         "HK"),
    "0405.HK": ("Yuexiu REIT",           "HK"),
    "0006.HK": ("Power Assets",          "HK"),
    "0002.HK": ("CLP Holdings",          "HK"),
    "0003.HK": ("HK & China Gas",        "HK"),
    "1038.HK": ("CK Infrastructure",     "HK"),
    "0177.HK": ("Jiangsu Expressway",    "HK"),
    "0548.HK": ("Shenzhen Expressway",   "HK"),
    "0728.HK": ("China Telecom",         "HK"),
    "0762.HK": ("China Unicom",          "HK"),
    "0941.HK": ("China Mobile",          "HK"),
    "2318.HK": ("Ping An Insurance",     "HK"),
    "0966.HK": ("China Taiping",         "HK"),
    "1336.HK": ("New China Life",        "HK"),
    "0857.HK": ("PetroChina",            "HK"),
    "0386.HK": ("Sinopec",              "HK"),
    "1088.HK": ("China Shenhua",         "HK"),
    "1171.HK": ("Yanzhou Coal",          "HK"),
    "0083.HK": ("Sino Land",             "HK"),
    "0001.HK": ("CK Hutchison",          "HK"),
    "0019.HK": ("Swire Pacific A",       "HK"),
    "0087.HK": ("Swire Pacific B",       "HK"),
    "0066.HK": ("MTR Corporation",       "HK"),
    "0151.HK": ("Want Want China",       "HK"),
    "0291.HK": ("China Resources Beer",  "HK"),
    "1929.HK": ("Chow Tai Fook",         "HK"),
    "0960.HK": ("Longfor Group",         "HK"),
    "2319.HK": ("Mengniu Dairy",         "HK"),
    "6823.HK": ("HKT Trust",             "HK"),
    "0270.HK": ("Guangdong Investment",  "HK"),
    "0659.HK": ("NWS Holdings",          "HK"),
    "1997.HK": ("Wharf REIC",            "HK"),
    "0694.HK": ("Beijing Enterprises",   "HK"),
    "0836.HK": ("China Resources Power", "HK"),
    "0358.HK": ("Jiangxi Copper",        "HK"),
    "2600.HK": ("Aluminum Corp China",   "HK"),
    "1113.HK": ("CR Land",               "HK"),
    # ── CN（A股 — 滬深300高息藍籌）──────────────────────────
    "600036.SS": ("招商銀行",            "CN"),
    "601398.SS": ("工商銀行",            "CN"),
    "601288.SS": ("農業銀行",            "CN"),
    "601939.SS": ("建設銀行",            "CN"),
    "601988.SS": ("中國銀行",            "CN"),
    "601328.SS": ("交通銀行",            "CN"),
    "600016.SS": ("民生銀行",            "CN"),
    "601166.SS": ("興業銀行",            "CN"),
    "600028.SS": ("中國石化",            "CN"),
    "601857.SS": ("中國石油",            "CN"),
    "600900.SS": ("長江電力",            "CN"),
    "601088.SS": ("中國神華",            "CN"),
    "600011.SS": ("華能國際",            "CN"),
    "601601.SS": ("中國太保",            "CN"),
    "601318.SS": ("中國平安",            "CN"),
    "601336.SS": ("新華保險",            "CN"),
    "600050.SS": ("中國聯通",            "CN"),
    "600886.SS": ("國投電力",            "CN"),
    "601186.SS": ("中國鐵建",            "CN"),
    "601800.SS": ("中國交建",            "CN"),
    "601390.SS": ("中國中鐵",            "CN"),
    "601668.SS": ("中國建築",            "CN"),
    "601225.SS": ("陝西煤業",            "CN"),
    "601666.SS": ("平煤股份",            "CN"),
    "600188.SS": ("兗礦能源",            "CN"),
    "600519.SS": ("貴州茅台",            "CN"),
    "000858.SZ": ("五糧液",              "CN"),
    "600887.SS": ("伊利股份",            "CN"),
    "000333.SZ": ("美的集團",            "CN"),
    "000651.SZ": ("格力電器",            "CN"),
    # ── US ────────────────────────────────────────────────
    "O":    ("Realty Income",            "US"),
    "AMT":  ("American Tower",           "US"),
    "PLD":  ("Prologis",                 "US"),
    "SPG":  ("Simon Property",           "US"),
    "VTR":  ("Ventas",                   "US"),
    "WELL": ("Welltower",                "US"),
    "NNN":  ("NNN REIT",                 "US"),
    "STAG": ("STAG Industrial",          "US"),
    "VICI": ("VICI Properties",          "US"),
    "NEE":  ("NextEra Energy",           "US"),
    "DUK":  ("Duke Energy",              "US"),
    "SO":   ("Southern Company",         "US"),
    "D":    ("Dominion Energy",          "US"),
    "AEP":  ("American Elec Power",      "US"),
    "XEL":  ("Xcel Energy",              "US"),
    "WEC":  ("WEC Energy",               "US"),
    "ES":   ("Eversource Energy",        "US"),
    "JPM":  ("JPMorgan Chase",           "US"),
    "BAC":  ("Bank of America",          "US"),
    "WFC":  ("Wells Fargo",              "US"),
    "C":    ("Citigroup",                "US"),
    "USB":  ("US Bancorp",               "US"),
    "PRU":  ("Prudential Financial",     "US"),
    "MET":  ("MetLife",                  "US"),
    "XOM":  ("ExxonMobil",              "US"),
    "CVX":  ("Chevron",                  "US"),
    "COP":  ("ConocoPhillips",           "US"),
    "EOG":  ("EOG Resources",            "US"),
    "PSX":  ("Phillips 66",              "US"),
    "JNJ":  ("Johnson & Johnson",        "US"),
    "ABBV": ("AbbVie",                   "US"),
    "MRK":  ("Merck",                    "US"),
    "PFE":  ("Pfizer",                   "US"),
    "KO":   ("Coca-Cola",                "US"),
    "PEP":  ("PepsiCo",                  "US"),
    "PG":   ("Procter & Gamble",         "US"),
    "MO":   ("Altria Group",             "US"),
    "PM":   ("Philip Morris",            "US"),
    "T":    ("AT&T",                     "US"),
    "VZ":   ("Verizon",                  "US"),
    "SCHD": ("Schwab US Dividend ETF",   "US"),
    "VYM":  ("Vanguard High Div ETF",    "US"),
    "HDV":  ("iShares Core Dividend",    "US"),
    "JEPI": ("JPM Equity Premium Inc",   "US"),
    "JEPQ": ("JPM Nasdaq Eq Premium",    "US"),
}

def main():
    out_path = "tickers_config.json"

    # 轉換格式：{ticker: {name, market, added_date, source}}
    data = {
        "_meta": {
            "version":      "1.0",
            "last_updated": str(datetime.date.today()),
            "description":  "股息追蹤清單 — 由 batch_importer 使用",
        },
        "tickers": {}
    }
    for ticker, (name, market) in TICKERS_CONFIG.items():
        data["tickers"][ticker] = {
            "name":       name,
            "market":     market,
            "added_date": "2024-01-01",   # 原有股票設為起始日
            "source":     "initial",       # initial / screener / manual
        }

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    total = len(data["tickers"])
    by_mkt = {}
    for v in data["tickers"].values():
        by_mkt[v["market"]] = by_mkt.get(v["market"], 0) + 1

    print(f"✅ 匯出完成：{out_path}")
    print(f"   總計 {total} 隻股票")
    for mkt, cnt in sorted(by_mkt.items()):
        print(f"   {mkt}: {cnt} 隻")

if __name__ == "__main__":
    main()
