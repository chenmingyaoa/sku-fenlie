import json
import re
import statistics
from collections import Counter
from dataclasses import dataclass, field
from html import unescape
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
from urllib.parse import quote

import requests

from xlsx_toolkit import WorkbookData


STOP_WORDS = {
    "a",
    "an",
    "and",
    "for",
    "the",
    "with",
    "to",
    "of",
    "in",
    "on",
    "by",
    "new",
    "uk",
    "usb",
    "pack",
    "pcs",
    "pc",
    "set",
    "adapter",
    "plug",
    "cable",
}

TURNOVER_ALIASES = {
    "sku": ["sku", "商品sku", "商品code编码", "商品code（编码）", "商品code(编码)"],
    "product_code": ["商品code编码", "商品code（编码）", "商品code(编码)"],
    "title": ["sku名称", "商品名称", "产品名称", "listingtitle", "标题"],
    "site": ["站点", "site"],
    "group": ["组别", "小组"],
    "department": ["部门"],
    "sales_30d": ["30天销量", "30天销量件数", "30day销量"],
    "turnover_rate": ["动销率"],
    "rotation_rate": ["轮动率"],
    "price": ["上架售价", "平均单价", "售价"],
    "cost": ["商品成本价", "成本价"],
    "inventory": ["现有库存", "总和库存"],
    "developer": ["开发员", "销售员"],
    "warehouse_country": ["仓库国家简码", "国家简码"],
}

TRAFFIC_ALIASES = {
    "listing_title": ["listingtitle", "title", "listing标题"],
    "item_id": ["ebayitemid", "itemid", "item id"],
    "category": ["category", "类目"],
    "total_impressions": ["totalimpressions"],
    "page_views": ["totalpageviews", "pageviews"],
    "quantity_sold": ["quantitysold"],
    "conversion_rate": ["salesconversionratequantitysoldtotalpageviews", "salesconversionrate"],
    "ctr": ["clickthroughratepageviewsfromebaysitetotalimpressions", "clickthroughrate"],
}

PAIN_POINT_RULES = [
    (
        {"adapter", "charger", "power", "usb", "typec", "pd", "dc"},
        [
            "买家通常担心兼容性、接口规格和供电稳定性。",
            "建议强调电压、电流、插头制式、适配设备和安全保护说明。",
        ],
    ),
    (
        {"bike", "bicycle", "motorcycle", "phone", "mount", "cycling", "glove"},
        [
            "买家更关注安装稳固、防滑、防震和通用尺寸。",
            "建议突出适配范围、安装场景、抗震防摔和户外耐用性。",
        ],
    ),
    (
        {"hearing", "aid", "filter", "wax"},
        [
            "买家通常在意适配型号、卫生清洁效果和更换便利度。",
            "建议补充兼容机型、单包数量、安装步骤和使用周期说明。",
        ],
    ),
]


@dataclass
class AnalysisOptions:
    group_filters: List[str] = field(default_factory=lambda: ["__all__"])
    site_filter: str = "__all__"
    min_sales_30d: float = 3.0
    min_turnover_rate: float = 10.0
    low_impressions_threshold: float = 100.0
    low_page_views_threshold: float = 10.0
    low_conversion_threshold: float = 0.02
    max_results: int = 30
    enable_ebay_scrape: bool = True
    enable_openai: bool = False
    openai_api_key: str = ""
    openai_model: str = "gpt-4.1-mini"
    ebay_site_domain: str = "www.ebay.co.uk"


def analyze_store_data(
    turnover_workbook: WorkbookData,
    traffic_workbook: WorkbookData,
    options: AnalysisOptions,
) -> Dict[str, object]:
    turnover_table = extract_table(turnover_workbook, TURNOVER_ALIASES, "动销")
    traffic_table = extract_table(traffic_workbook, TRAFFIC_ALIASES, "流量")

    turnover_rows = normalize_rows(turnover_table["rows"], turnover_table["field_map"])
    traffic_rows = normalize_rows(traffic_table["rows"], traffic_table["field_map"])
    traffic_records = aggregate_traffic_records(traffic_rows)
    market_rules = build_market_rules(traffic_records, options)

    filtered_turnover = [row for row in turnover_rows if turnover_row_matches(row, options)]
    traffic_index = build_traffic_index(traffic_records)

    candidates = []
    for turnover_row in filtered_turnover:
        evaluated = evaluate_candidate(turnover_row, traffic_index, market_rules, options)
        if evaluated is not None:
            candidates.append(evaluated)

    candidates.sort(
        key=lambda row: (
            row["conversion_sort"],
            -row["opportunity_score"],
            -row["sales_30d_num"],
        )
    )
    candidates = candidates[: max(1, options.max_results)]

    enriched_rows = []
    for rank, candidate in enumerate(candidates, start=1):
        candidate["rank"] = rank
        candidate.update(build_market_snapshot(candidate, options))
        enriched_rows.append(candidate)

    summary = build_summary(filtered_turnover, traffic_records, enriched_rows, turnover_table, traffic_table, market_rules)
    return {
        "summary": summary,
        "rows": enriched_rows,
        "export_rows": build_export_rows(enriched_rows),
        "turnover_sheet": turnover_table["sheet_name"],
        "traffic_sheet": traffic_table["sheet_name"],
        "market_rules": market_rules,
    }


def extract_turnover_filter_options(workbook: WorkbookData) -> Dict[str, List[str]]:
    turnover_table = extract_table(workbook, TURNOVER_ALIASES, "动销")
    turnover_rows = normalize_rows(turnover_table["rows"], turnover_table["field_map"])
    groups = sorted({row.get("group", "").strip() for row in turnover_rows if row.get("group", "").strip()})
    sites = sorted({row.get("site", "").strip() for row in turnover_rows if row.get("site", "").strip()})
    return {"groups": groups, "sites": sites}


def extract_table(workbook: WorkbookData, aliases: Dict[str, Sequence[str]], label: str) -> Dict[str, object]:
    best_result = None
    best_score = -1
    for sheet in workbook.sheets:
        header_index, header_score = find_header_row(sheet.rows, aliases)
        if header_index is None or header_score < best_score:
            continue
        headers = list(sheet.rows[header_index])
        field_map = map_headers(headers, aliases)
        if not field_map:
            continue
        best_score = header_score
        best_result = {
            "sheet_name": sheet.name,
            "headers": headers,
            "rows": [pad_row(row, len(headers)) for row in sheet.rows[header_index + 1 :]],
            "field_map": field_map,
        }
    if not best_result:
        raise ValueError(f"{label}文件里没有识别到可分析的数据表，请确认表头是否完整。")
    return best_result


def find_header_row(rows: Sequence[Sequence[str]], aliases: Dict[str, Sequence[str]]) -> Tuple[Optional[int], int]:
    best_index = None
    best_score = -1
    for index, row in enumerate(rows[:12]):
        headers = [clean_header_name(cell) for cell in row]
        score = 0
        for alias_list in aliases.values():
            if any(alias in headers for alias in alias_list):
                score += 1
        if score > best_score:
            best_score = score
            best_index = index
    return best_index, best_score


def map_headers(headers: Sequence[str], aliases: Dict[str, Sequence[str]]) -> Dict[str, int]:
    normalized = [clean_header_name(header) for header in headers]
    field_map: Dict[str, int] = {}
    for field, alias_list in aliases.items():
        for alias in alias_list:
            if alias in normalized:
                field_map[field] = normalized.index(alias)
                break
    return field_map


def normalize_rows(rows: Sequence[Sequence[str]], field_map: Dict[str, int]) -> List[Dict[str, str]]:
    records = []
    for raw_row in rows:
        if not any(str(value).strip() for value in raw_row):
            continue
        record = {field: safe_get(raw_row, index) for field, index in field_map.items()}
        if any(record.values()):
            records.append(record)
    return records


def aggregate_traffic_records(rows: Sequence[Dict[str, str]]) -> List[Dict[str, object]]:
    grouped: Dict[Tuple[str, str], Dict[str, object]] = {}
    for row in rows:
        item_id = (row.get("item_id") or "").strip()
        listing_title = (row.get("listing_title") or "").strip()
        if not item_id and not listing_title:
            continue
        key = (item_id, listing_title)
        bucket = grouped.setdefault(
            key,
            {
                "item_id": item_id,
                "listing_title": listing_title,
                "category": row.get("category", ""),
                "impressions": 0.0,
                "page_views": 0.0,
                "quantity_sold": 0.0,
                "ctr_values": [],
                "conversion_values": [],
                "record_count": 0,
            },
        )
        bucket["impressions"] += parse_number(row.get("total_impressions")) or 0.0
        bucket["page_views"] += parse_number(row.get("page_views")) or 0.0
        bucket["quantity_sold"] += parse_number(row.get("quantity_sold")) or 0.0
        ctr = normalize_ratio(row.get("ctr"))
        conversion = normalize_ratio(row.get("conversion_rate"))
        if ctr is not None:
            bucket["ctr_values"].append(ctr)
        if conversion is not None:
            bucket["conversion_values"].append(conversion)
        bucket["record_count"] += 1

    records = []
    for (_item_id, _title), bucket in grouped.items():
        page_views = bucket["page_views"]
        quantity_sold = bucket["quantity_sold"]
        records.append(
            {
                "item_id": bucket["item_id"],
                "listing_title": bucket["listing_title"],
                "traffic_key": build_traffic_key(bucket["item_id"], bucket["listing_title"]),
                "category": bucket["category"],
                "impressions": bucket["impressions"] or None,
                "page_views": bucket["page_views"] or None,
                "quantity_sold": bucket["quantity_sold"] or None,
                "ctr": mean_or_none(bucket["ctr_values"]),
                "conversion_rate": (quantity_sold / page_views) if page_views else mean_or_none(bucket["conversion_values"]),
                "record_count": bucket["record_count"],
            }
        )
    return records


def build_market_rules(traffic_records: Sequence[Dict[str, object]], options: AnalysisOptions) -> Dict[str, object]:
    impressions = positive_numbers(record.get("impressions") for record in traffic_records)
    page_views = positive_numbers(record.get("page_views") for record in traffic_records)
    ctr_values = positive_numbers(record.get("ctr") for record in traffic_records)
    conversion_values = positive_numbers(record.get("conversion_rate") for record in traffic_records)

    impression_floor = max(options.low_impressions_threshold, percentile(impressions, 25) or options.low_impressions_threshold)
    page_view_floor = max(options.low_page_views_threshold, percentile(page_views, 25) or options.low_page_views_threshold)
    ctr_floor = min(max(0.003, percentile(ctr_values, 25) or 0.003), 0.02)
    conversion_floor = min(max(options.low_conversion_threshold, percentile(conversion_values, 25) or options.low_conversion_threshold), 0.05)

    return {
        "impression_floor": round(impression_floor, 2),
        "page_view_floor": round(page_view_floor, 2),
        "ctr_floor": round(ctr_floor, 4),
        "conversion_floor": round(conversion_floor, 4),
        "rule_lines": [
            f"曝光弱：Total impressions < {round(impression_floor, 2)}",
            f"点击弱：Total page views < {round(page_view_floor, 2)} 或 CTR < {round(ctr_floor * 100, 2)}%",
            f"转化弱：Sales conversion rate < {round(conversion_floor * 100, 2)}%",
            "流量主键：Item ID + Listing title 聚合后判定",
        ],
    }


def turnover_row_matches(row: Dict[str, str], options: AnalysisOptions) -> bool:
    selected_groups = [value for value in options.group_filters if value and value != "__all__"]
    if selected_groups and row.get("group", "").strip() not in selected_groups:
        return False
    if options.site_filter != "__all__" and row.get("site", "").strip() != options.site_filter:
        return False
    return True


def build_traffic_index(traffic_records: Sequence[Dict[str, object]]) -> Dict[str, object]:
    index = {"records": list(traffic_records), "titles": {}, "search_blob": []}
    for record in traffic_records:
        title_key = normalize_text(str(record.get("listing_title", "")))
        if title_key:
            index["titles"].setdefault(title_key, []).append(record)
        blob = " ".join([str(record.get("listing_title", "")), str(record.get("item_id", ""))]).lower()
        index["search_blob"].append((blob, record))
    return index


def evaluate_candidate(
    turnover_row: Dict[str, str],
    traffic_index: Dict[str, object],
    market_rules: Dict[str, object],
    options: AnalysisOptions,
) -> Optional[Dict[str, object]]:
    sales_30d = parse_number(turnover_row.get("sales_30d"))
    turnover_rate = normalize_ratio(turnover_row.get("turnover_rate"))
    good_sales = sales_30d is not None and sales_30d >= options.min_sales_30d
    good_turnover = turnover_rate is not None and turnover_rate >= options.min_turnover_rate
    if not (good_sales or good_turnover):
        return None

    matched_record, match_reason, match_score = match_traffic_record(turnover_row, traffic_index)
    issue_type, issue_reasons, include_row = classify_issue(matched_record, market_rules)
    if not include_row:
        return None

    turnover_title = turnover_row.get("title") or turnover_row.get("sku") or ""
    traffic_title = str(matched_record.get("listing_title", "")) if matched_record else ""
    own_price = parse_number(turnover_row.get("price"))

    impressions = get_float(matched_record, "impressions")
    page_views = get_float(matched_record, "page_views")
    quantity_sold = get_float(matched_record, "quantity_sold")
    ctr = get_float(matched_record, "ctr")
    conversion = get_float(matched_record, "conversion_rate")

    return {
        "sku": turnover_row.get("sku") or turnover_row.get("product_code") or "",
        "title": turnover_title,
        "traffic_title": traffic_title,
        "site": turnover_row.get("site", ""),
        "group": turnover_row.get("group", ""),
        "sales_30d": format_number(sales_30d),
        "sales_30d_num": sales_30d or 0.0,
        "turnover_rate": format_percent(turnover_rate),
        "turnover_rate_num": turnover_rate or 0.0,
        "own_price": format_currency(own_price),
        "own_price_num": own_price,
        "inventory": turnover_row.get("inventory", ""),
        "item_id": str(matched_record.get("item_id", "")) if matched_record else "",
        "traffic_key": str(matched_record.get("traffic_key", "")) if matched_record else "",
        "impressions": format_number(impressions),
        "impressions_num": impressions,
        "page_views": format_number(page_views),
        "page_views_num": page_views,
        "quantity_sold": format_number(quantity_sold),
        "quantity_sold_num": quantity_sold,
        "ctr": format_percent(ctr),
        "ctr_num": ctr,
        "conversion_rate": format_percent(conversion),
        "conversion_rate_num": conversion,
        "issue_type": issue_type,
        "issue_reasons": "；".join(issue_reasons),
        "match_reason": match_reason,
        "match_score": match_score,
        "conversion_sort": conversion if conversion is not None else -1.0,
        "opportunity_score": round(
            calculate_opportunity_score(
                sales_30d=sales_30d or 0.0,
                turnover_rate=turnover_rate or 0.0,
                impressions=impressions,
                page_views=page_views,
                ctr=ctr,
                conversion=conversion,
                has_match=matched_record is not None,
                market_rules=market_rules,
            ),
            2,
        ),
        "keyword_seed": build_keyword_seed(traffic_title or turnover_title, turnover_title),
    }


def match_traffic_record(
    turnover_row: Dict[str, str],
    traffic_index: Dict[str, object],
) -> Tuple[Optional[Dict[str, object]], str, float]:
    sku = (turnover_row.get("sku") or "").strip().lower()
    title = turnover_row.get("title") or ""
    normalized_title = normalize_text(title)

    exact_matches = traffic_index["titles"].get(normalized_title, [])
    if exact_matches:
        return exact_matches[0], "标题精确匹配", 1.0

    if sku:
        for blob, record in traffic_index["search_blob"]:
            if sku in blob:
                return record, "SKU 或 Item ID 命中", 0.92

    turnover_tokens = tokenize(title)
    if not turnover_tokens:
        return None, "未匹配到流量主键", 0.0

    scored = []
    for record in traffic_index["records"]:
        traffic_tokens = tokenize(str(record.get("listing_title", "")))
        overlap = turnover_tokens & traffic_tokens
        if not overlap:
            continue
        score = len(overlap) / max(len(turnover_tokens), len(traffic_tokens))
        if score >= 0.34:
            scored.append((score, record))

    scored.sort(key=lambda item: item[0], reverse=True)
    if scored:
        best_score, best_record = scored[0]
        return best_record, "标题关键词模糊匹配", round(best_score, 2)
    return None, "未匹配到流量主键", 0.0


def classify_issue(
    matched_record: Optional[Dict[str, object]],
    market_rules: Dict[str, object],
) -> Tuple[str, List[str], bool]:
    if matched_record is None:
        return "无流量数据", ["未匹配到 Item ID + Listing title 流量主键"], True

    reasons = []
    issue_bucket = []
    impressions = matched_record.get("impressions")
    page_views = matched_record.get("page_views")
    ctr = matched_record.get("ctr")
    conversion = matched_record.get("conversion_rate")

    if impressions is None or impressions < market_rules["impression_floor"]:
        issue_bucket.append("流量偏低")
        reasons.append(f"曝光低于阈值 {market_rules['impression_floor']}")
    if page_views is None or page_views < market_rules["page_view_floor"]:
        issue_bucket.append("流量偏低")
        reasons.append(f"浏览低于阈值 {market_rules['page_view_floor']}")
    if ctr is not None and ctr < market_rules["ctr_floor"]:
        issue_bucket.append("流量偏低")
        reasons.append(f"CTR 低于阈值 {round(market_rules['ctr_floor'] * 100, 2)}%")
    if conversion is None or conversion < market_rules["conversion_floor"]:
        issue_bucket.append("转化偏低")
        reasons.append(f"转化率低于阈值 {round(market_rules['conversion_floor'] * 100, 2)}%")

    if issue_bucket:
        if "流量偏低" in issue_bucket:
            return "流量偏低", reasons, True
        return "转化偏低", reasons, True
    return "表现正常", [], False


def calculate_opportunity_score(
    sales_30d: float,
    turnover_rate: float,
    impressions: Optional[float],
    page_views: Optional[float],
    ctr: Optional[float],
    conversion: Optional[float],
    has_match: bool,
    market_rules: Dict[str, object],
) -> float:
    sales_component = min(40.0, sales_30d * 2.2)
    turnover_component = min(24.0, turnover_rate * 0.4)
    no_match_bonus = 22.0 if not has_match else 0.0
    exposure_gap = 0.0 if impressions is None else max(0.0, market_rules["impression_floor"] - impressions) / max(market_rules["impression_floor"], 1.0)
    click_gap = 0.0 if page_views is None else max(0.0, market_rules["page_view_floor"] - page_views) / max(market_rules["page_view_floor"], 1.0)
    ctr_gap = 0.0 if ctr is None else max(0.0, market_rules["ctr_floor"] - ctr) / max(market_rules["ctr_floor"], 0.0001)
    conversion_gap = 0.0 if conversion is None else max(0.0, market_rules["conversion_floor"] - conversion) / max(market_rules["conversion_floor"], 0.0001)
    return sales_component + turnover_component + no_match_bonus + exposure_gap * 10 + click_gap * 10 + ctr_gap * 8 + conversion_gap * 12


def build_market_snapshot(candidate: Dict[str, object], options: AnalysisOptions) -> Dict[str, object]:
    source_title = str(candidate.get("traffic_title") or candidate.get("title") or "")
    search_keywords = extract_search_keywords(source_title)
    search_query = compact_query(search_keywords or source_title)
    if not search_query:
        return empty_market_snapshot(candidate)

    active_items: List[Dict[str, object]] = []
    sold_items: List[Dict[str, object]] = []
    if options.enable_ebay_scrape:
        active_items = search_ebay_active(search_query, options.ebay_site_domain)
        sold_items = search_ebay_sold(search_query, options.ebay_site_domain)

    active_titles = [item["title"] for item in active_items]
    sold_titles = [item["title"] for item in sold_items]
    active_price_summary = summarize_prices(active_items, "未抓到在售竞品价格")
    sold_price_summary = summarize_prices(sold_items, "未抓到已售出竞品价格")
    common_keywords = extract_common_keywords(candidate["keyword_seed"], active_titles + sold_titles)
    active_scores = score_competitors(active_items, candidate.get("own_price_num"), search_query)
    sold_scores = score_competitors(sold_items, candidate.get("own_price_num"), search_query)
    discount_summary = summarize_discount_tactics(active_items)
    shipping_summary = summarize_shipping_messages(active_items)

    heuristic_advice = build_heuristic_advice(
        candidate,
        common_keywords,
        active_price_summary,
        sold_price_summary,
        active_scores,
        discount_summary,
        shipping_summary,
    )
    ai_summary = None
    if options.enable_openai and options.openai_api_key.strip():
        ai_summary = run_openai_market_analysis(
            candidate,
            common_keywords,
            active_price_summary,
            sold_price_summary,
            active_scores,
            sold_scores,
            discount_summary,
            shipping_summary,
            options,
        )
    merged = merge_ai_with_heuristic(ai_summary, heuristic_advice)

    return {
        "market_query_title": search_query,
        "market_query_keywords": search_keywords,
        "active_sample_count": len(active_items),
        "sold_sample_count": len(sold_items),
        "active_price_range": active_price_summary["range_text"],
        "price_range": sold_price_summary["range_text"],
        "common_keywords": ", ".join(common_keywords[:8]),
        "discount_summary": discount_summary,
        "shipping_summary": shipping_summary,
        "competitor_score_text": "; ".join(
            f'{item["score"]}/100 | {item["price_text"] or "-"} | {item["discount_text"] or item["shipping_text"] or "-"} | {item["title"][:38]}'
            for item in active_scores[:3]
        ),
        "competitor_links": "\n".join(item["link"] for item in active_scores[:5] if item["link"]),
        "market_summary": merged["market_summary"],
        "title_suggestion": merged["title_suggestion"],
        "listing_advice": merged["listing_advice"],
        "pain_points": merged["pain_points"],
    }


def empty_market_snapshot(candidate: Dict[str, object]) -> Dict[str, object]:
    title = str(candidate.get("traffic_title") or candidate.get("title") or "")
    return {
        "market_query_title": title,
        "market_query_keywords": "",
        "active_sample_count": 0,
        "sold_sample_count": 0,
        "active_price_range": "没有可用于竞品检索的标题",
        "price_range": "没有可用于竞品检索的标题",
        "common_keywords": "",
        "discount_summary": "暂无折扣手段数据",
        "shipping_summary": "暂无英国配送信息",
        "competitor_score_text": "",
        "competitor_links": "",
        "market_summary": "没有可用于竞品检索的标题。",
        "title_suggestion": "请先确认流量标题或动销标题是否完整。",
        "listing_advice": "建议补足标题后再做竞品页面分析。",
        "pain_points": "当前数据不足，无法稳定提炼痛点卖点。",
    }


def search_ebay_active(query: str, site_domain: str, limit: int = 12) -> List[Dict[str, object]]:
    search_url = f"https://{site_domain}/sch/i.html?_nkw={quote(query)}&_ipg={limit}&rt=nc&_sop=12"
    return fetch_ebay_results(search_url)


def search_ebay_sold(query: str, site_domain: str, limit: int = 12) -> List[Dict[str, object]]:
    search_url = f"https://{site_domain}/sch/i.html?_nkw={quote(query)}&_ipg={limit}&LH_Sold=1&LH_Complete=1&rt=nc&_sop=13"
    return fetch_ebay_results(search_url)


def fetch_ebay_results(search_url: str) -> List[Dict[str, object]]:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
        ),
        "Accept-Language": "en-GB,en;q=0.9",
        "Referer": "https://www.ebay.co.uk/",
    }
    try:
        response = requests.get(search_url, headers=headers, timeout=15)
        response.raise_for_status()
    except Exception:
        return []

    html = response.text
    blocks = re.split(r'<li[^>]+class="[^"]*s-item[^"]*"', html)
    results = []
    for block in blocks[1:]:
        parsed = parse_ebay_result_block(block)
        if parsed:
            results.append(parsed)
        if len(results) >= 12:
            break
    return results


def parse_ebay_result_block(block: str) -> Optional[Dict[str, object]]:
    link_match = re.search(r'<a[^>]+class="[^"]*s-item__link[^"]*"[^>]+href="([^"]+)"', block, re.I)
    title_match = re.search(
        r'<(?:div|span)[^>]+class="[^"]*s-item__title[^"]*"[^>]*>(.*?)</(?:div|span)>',
        block,
        re.I | re.S,
    )
    if not link_match or not title_match:
        return None

    title = strip_tags(title_match.group(1))
    if not title or "Shop on eBay" in title:
        return None

    price_text = extract_first_match(block, r'<span[^>]+class="[^"]*s-item__price[^"]*"[^>]*>(.*?)</span>')
    shipping_text = extract_first_match(block, r'<span[^>]+class="[^"]*s-item__shipping[^"]*"[^>]*>(.*?)</span>')
    sold_hint = extract_first_match(block, r'<span[^>]+class="[^"]*POSITIVE[^"]*"[^>]*>(.*?)</span>')
    condition_text = extract_first_match(block, r'<span[^>]+class="[^"]*SECONDARY_INFO[^"]*"[^>]*>(.*?)</span>')
    seller_text = extract_first_match(block, r'<span[^>]+class="[^"]*s-item__seller-info-text[^"]*"[^>]*>(.*?)</span>')

    return {
        "title": title,
        "link": html_unescape(link_match.group(1)),
        "price_text": strip_tags(price_text) if price_text else "",
        "price_values": parse_price_candidates(strip_tags(price_text) if price_text else ""),
        "shipping_text": strip_tags(shipping_text) if shipping_text else "",
        "sold_hint": strip_tags(sold_hint) if sold_hint else "",
        "condition_text": strip_tags(condition_text) if condition_text else "",
        "seller_text": strip_tags(seller_text) if seller_text else "",
        "discount_text": extract_discount_text(block),
    }


def score_competitors(competitors: Sequence[Dict[str, object]], own_price: Optional[float], search_title: str) -> List[Dict[str, object]]:
    own_tokens = tokenize(search_title)
    results = []
    for competitor in competitors:
        title_tokens = tokenize(competitor["title"])
        title_overlap = len(own_tokens & title_tokens) / max(len(own_tokens), len(title_tokens)) if own_tokens and title_tokens else 0.0
        price_values = competitor.get("price_values") or []
        competitor_price = statistics.mean(price_values) if price_values else None
        if own_price and competitor_price:
            price_gap = abs(own_price - competitor_price) / max(own_price, competitor_price, 1.0)
            price_score = max(0.0, 1.0 - price_gap)
        else:
            price_score = 0.45
        enriched = dict(competitor)
        enriched["score"] = round(title_overlap * 70 + price_score * 30)
        results.append(enriched)
    results.sort(key=lambda item: item["score"], reverse=True)
    return results


def summarize_prices(competitors: Sequence[Dict[str, object]], empty_text: str) -> Dict[str, object]:
    values = []
    for competitor in competitors:
        values.extend(competitor.get("price_values") or [])
    if not values:
        return {"range_text": empty_text, "median": None, "min": None, "max": None}
    min_value = min(values)
    max_value = max(values)
    median_value = statistics.median(values)
    return {
        "range_text": f"GBP {min_value:.2f} - {max_value:.2f}，中位价约 GBP {median_value:.2f}",
        "median": median_value,
        "min": min_value,
        "max": max_value,
    }


def summarize_discount_tactics(active_items: Sequence[Dict[str, object]]) -> str:
    phrases = Counter()
    for item in active_items:
        for text in [item.get("discount_text", ""), item.get("sold_hint", "")]:
            cleaned = str(text).strip()
            if cleaned:
                phrases[cleaned] += 1
    if not phrases:
        return "未明显抓到折扣手段，可重点优化标题、主图和运费表达。"
    return "；".join(f"{text}（{count}）" for text, count in phrases.most_common(3))


def summarize_shipping_messages(active_items: Sequence[Dict[str, object]]) -> str:
    phrases = Counter()
    for item in active_items:
        cleaned = str(item.get("shipping_text", "")).strip()
        if cleaned:
            phrases[cleaned] += 1
    if not phrases:
        return "未明显抓到英国配送卖点。"
    return "；".join(f"{text}（{count}）" for text, count in phrases.most_common(3))


def build_heuristic_advice(
    candidate: Dict[str, object],
    common_keywords: Sequence[str],
    active_price_summary: Dict[str, object],
    sold_price_summary: Dict[str, object],
    active_scores: Sequence[Dict[str, object]],
    discount_summary: str,
    shipping_summary: str,
) -> Dict[str, str]:
    search_title = str(candidate.get("traffic_title") or candidate.get("title") or "")
    own_price = candidate.get("own_price_num")
    issue_type = candidate["issue_type"]
    missing_keywords = [keyword for keyword in common_keywords if keyword.lower() not in search_title.lower()][:4]
    keyword_tip = "、".join(missing_keywords) if missing_keywords else "核心兼容词、规格词、场景词"

    pricing_tip = "先补足流量，再观察价格段是否需要调整。"
    sold_median = sold_price_summary.get("median")
    active_median = active_price_summary.get("median")
    benchmark = sold_median if sold_median is not None else active_median
    if own_price and benchmark is not None:
        if own_price > benchmark * 1.12:
            pricing_tip = "当前售价高于市场中位价，建议先补强卖点表达，再测试轻度降价。"
        elif own_price < benchmark * 0.88:
            pricing_tip = "当前售价低于市场中位价，可以补强标题和主图卖点争取更高转化。"
        else:
            pricing_tip = "当前售价接近市场中位价，更适合优先优化关键词与卖点表达。"

    top_titles = ", ".join(item["title"][:28] for item in active_scores[:2]) or "未抓到足够在售竞品"
    return {
        "market_summary": f"这条链接当前诊断为“{issue_type}”。在售竞品主要集中在：{top_titles}。折扣手段参考：{discount_summary}",
        "title_suggestion": f"建议把标题补齐 {keyword_tip}，并把规格、适配范围、用途词放在前半段。",
        "listing_advice": f"{pricing_tip} 同时建议参考英国配送表达：{shipping_summary}",
        "pain_points": build_pain_points(search_title),
    }


def build_pain_points(title: str) -> str:
    title_tokens = tokenize(title)
    for keywords, advice_lines in PAIN_POINT_RULES:
        if title_tokens & keywords:
            return " ".join(advice_lines)
    return "建议围绕兼容性、规格参数、使用场景、质量保证和发货时效来组织卖点。"


def run_openai_market_analysis(
    candidate: Dict[str, object],
    common_keywords: Sequence[str],
    active_price_summary: Dict[str, object],
    sold_price_summary: Dict[str, object],
    active_scores: Sequence[Dict[str, object]],
    sold_scores: Sequence[Dict[str, object]],
    discount_summary: str,
    shipping_summary: str,
    options: AnalysisOptions,
) -> Optional[Dict[str, str]]:
    prompt = {
        "sku": candidate["sku"],
        "item_id": candidate["item_id"],
        "traffic_key": candidate["traffic_key"],
        "listing_title": candidate["traffic_title"],
        "turnover_title": candidate["title"],
        "issue_type": candidate["issue_type"],
        "issue_reasons": candidate["issue_reasons"],
        "search_keywords": extract_search_keywords(candidate.get("traffic_title") or candidate.get("title") or ""),
        "active_price_range": active_price_summary["range_text"],
        "sold_price_range": sold_price_summary["range_text"],
        "common_keywords": list(common_keywords[:10]),
        "discount_summary": discount_summary,
        "shipping_summary": shipping_summary,
        "active_competitors": active_scores[:5],
        "sold_competitors": sold_scores[:5],
    }
    system_prompt = (
        "你是 eBay 英国站运营分析师。"
        "请基于输入数据输出 JSON 对象，字段必须只有 market_summary、title_suggestion、listing_advice、pain_points。"
    )
    response = call_openai_chat_completions(system_prompt, json.dumps(prompt, ensure_ascii=False), options)
    if response:
        return response
    return call_openai_responses(system_prompt, json.dumps(prompt, ensure_ascii=False), options)


def call_openai_chat_completions(system_prompt: str, user_prompt: str, options: AnalysisOptions) -> Optional[Dict[str, str]]:
    try:
        response = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {options.openai_api_key.strip()}",
                "Content-Type": "application/json",
            },
            json={
                "model": options.openai_model.strip() or "gpt-4.1-mini",
                "temperature": 0.2,
                "response_format": {"type": "json_object"},
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
            },
            timeout=20,
        )
        response.raise_for_status()
        payload = response.json()["choices"][0]["message"]["content"]
        return normalize_ai_payload(json.loads(payload))
    except Exception:
        return None


def call_openai_responses(system_prompt: str, user_prompt: str, options: AnalysisOptions) -> Optional[Dict[str, str]]:
    try:
        response = requests.post(
            "https://api.openai.com/v1/responses",
            headers={
                "Authorization": f"Bearer {options.openai_api_key.strip()}",
                "Content-Type": "application/json",
            },
            json={
                "model": options.openai_model.strip() or "gpt-4.1-mini",
                "input": [
                    {"role": "system", "content": [{"type": "input_text", "text": system_prompt}]},
                    {"role": "user", "content": [{"type": "input_text", "text": user_prompt}]},
                ],
            },
            timeout=20,
        )
        response.raise_for_status()
        return normalize_ai_payload(json.loads(extract_response_text(response.json())))
    except Exception:
        return None


def normalize_ai_payload(payload: Dict[str, object]) -> Dict[str, str]:
    return {
        "market_summary": str(payload.get("market_summary", "")).strip(),
        "title_suggestion": str(payload.get("title_suggestion", "")).strip(),
        "listing_advice": str(payload.get("listing_advice", "")).strip(),
        "pain_points": str(payload.get("pain_points", "")).strip(),
    }


def extract_response_text(response_json: Dict[str, object]) -> str:
    for item in response_json.get("output", []):
        for content in item.get("content", []):
            if content.get("type") == "output_text":
                return content.get("text", "")
    return response_json.get("output_text", "")


def merge_ai_with_heuristic(ai_summary: Optional[Dict[str, str]], heuristic: Dict[str, str]) -> Dict[str, str]:
    if not ai_summary:
        return heuristic
    return {key: ai_summary.get(key) or heuristic[key] for key in heuristic}


def build_summary(
    turnover_rows: Sequence[Dict[str, str]],
    traffic_records: Sequence[Dict[str, object]],
    result_rows: Sequence[Dict[str, object]],
    turnover_table: Dict[str, object],
    traffic_table: Dict[str, object],
    market_rules: Dict[str, object],
) -> Dict[str, object]:
    return {
        "turnover_count": len(turnover_rows),
        "traffic_count": len(traffic_records),
        "candidate_count": len(result_rows),
        "no_traffic_count": sum(1 for row in result_rows if row["issue_type"] == "无流量数据"),
        "low_traffic_count": sum(1 for row in result_rows if row["issue_type"] == "流量偏低"),
        "low_conversion_count": sum(1 for row in result_rows if row["issue_type"] == "转化偏低"),
        "turnover_sheet": turnover_table["sheet_name"],
        "traffic_sheet": traffic_table["sheet_name"],
        "rule_lines": market_rules["rule_lines"],
    }


def build_export_rows(rows: Sequence[Dict[str, object]]) -> List[List[str]]:
    headers = [
        "优先级",
        "SKU",
        "Item ID",
        "流量唯一主键",
        "动销标题",
        "流量标题",
        "搜索关键词",
        "竞品搜索标题",
        "站点",
        "组别",
        "问题类型",
        "问题原因",
        "30天销量",
        "动销率",
        "当前售价",
        "Total impressions",
        "Total page views",
        "CTR",
        "Sales conversion rate",
        "在售样本数",
        "已售样本数",
        "在售价格区间",
        "已售价格区间",
        "折扣手段",
        "英国配送卖点",
        "公共关键词",
        "竞品评分",
        "市场总结",
        "标题建议",
        "上架建议",
        "痛点卖点建议",
        "竞品链接",
    ]
    export_rows = [headers]
    for row in rows:
        export_rows.append(
            [
                str(row["rank"]),
                str(row["sku"]),
                str(row["item_id"]),
                str(row["traffic_key"]),
                str(row["title"]),
                str(row["traffic_title"]),
                str(row["market_query_keywords"]),
                str(row["market_query_title"]),
                str(row["site"]),
                str(row["group"]),
                str(row["issue_type"]),
                str(row["issue_reasons"]),
                str(row["sales_30d"]),
                str(row["turnover_rate"]),
                str(row["own_price"]),
                str(row["impressions"]),
                str(row["page_views"]),
                str(row["ctr"]),
                str(row["conversion_rate"]),
                str(row["active_sample_count"]),
                str(row["sold_sample_count"]),
                str(row["active_price_range"]),
                str(row["price_range"]),
                str(row["discount_summary"]),
                str(row["shipping_summary"]),
                str(row["common_keywords"]),
                str(row["competitor_score_text"]),
                str(row["market_summary"]),
                str(row["title_suggestion"]),
                str(row["listing_advice"]),
                str(row["pain_points"]),
                str(row["competitor_links"]),
            ]
        )
    return export_rows


def build_traffic_key(item_id: str, listing_title: str) -> str:
    return f"{item_id} | {listing_title}".strip(" |")


def build_keyword_seed(*values: str) -> List[str]:
    counts = Counter()
    for value in values:
        counts.update(tokenize(value))
    return [token for token, _count in counts.most_common(10)]


def extract_common_keywords(seed_keywords: Sequence[str], competitor_titles: Sequence[str]) -> List[str]:
    counts = Counter(token.lower() for token in seed_keywords if token)
    for title in competitor_titles:
        counts.update(tokenize(title))
    return [token for token, _count in counts.most_common(10)]


def extract_search_keywords(title: str, max_terms: int = 10) -> str:
    tokens = ordered_tokens(title)
    cleaned = []
    seen = set()
    for token in tokens:
        normalized = token.lower().strip()
        if not normalized or normalized in STOP_WORDS or normalized in seen:
            continue
        seen.add(normalized)
        cleaned.append(token)
        if len(cleaned) >= max_terms:
            break
    return " ".join(cleaned)


def ordered_tokens(text: str) -> List[str]:
    if not text:
        return []
    tokens = re.findall(r"[A-Za-z0-9\-/\.]+|[\u4e00-\u9fff]{2,}", text)
    return [token.strip() for token in tokens if token.strip()]


def extract_discount_text(block: str) -> str:
    candidates = []
    patterns = [
        r'(Save up to[^<]+)',
        r'(Multi-buy[^<]+)',
        r'(Buy \d+ get \d+[^<]*)',
        r'(\d+% off[^<]*)',
        r'(discount[^<]*)',
    ]
    for pattern in patterns:
        match = re.search(pattern, block, re.I)
        if match:
            candidates.append(strip_tags(match.group(1)))
    return "；".join(dict.fromkeys(text for text in candidates if text))


def extract_first_match(text: str, pattern: str) -> str:
    match = re.search(pattern, text, re.I | re.S)
    return match.group(1) if match else ""


def tokenize(text: str) -> set:
    if not text:
        return set()
    chunks = re.findall(r"[A-Za-z0-9\-/\.]+|[\u4e00-\u9fff]{2,}", text.lower())
    tokens = set()
    for chunk in chunks:
        normalized = chunk.strip("-./ ").replace(" ", "")
        if not normalized or normalized in STOP_WORDS:
            continue
        if len(normalized) == 1 and not normalized.isdigit():
            continue
        tokens.add(normalized)
    return tokens


def normalize_text(text: str) -> str:
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", "", (text or "").lower())


def compact_query(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())[:120]


def clean_header_name(text: str) -> str:
    raw = (text or "").strip().lower()
    replacements = {"\n": "", "\r": "", " ": "", "（": "(", "）": ")", "￥": "¥", "＝": "="}
    for source, target in replacements.items():
        raw = raw.replace(source, target)
    return re.sub(r"[^\w\u4e00-\u9fff%=¥()]+", "", raw)


def safe_get(row: Sequence[str], index: int) -> str:
    return row[index].strip() if index < len(row) else ""


def pad_row(row: Sequence[str], width: int) -> List[str]:
    values = list(row)
    if len(values) < width:
        values.extend([""] * (width - len(values)))
    return values[:width]


def get_float(record: Optional[Dict[str, object]], key: str) -> Optional[float]:
    if not record:
        return None
    value = record.get(key)
    return float(value) if value is not None else None


def parse_number(value: Optional[str]) -> Optional[float]:
    text = str(value or "").strip().replace(",", "")
    if not text or text == "-":
        return None
    match = re.findall(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    try:
        return float(match[0])
    except ValueError:
        return None


def normalize_ratio(value: Optional[str]) -> Optional[float]:
    text = str(value or "").strip()
    if not text or text == "-":
        return None
    number = parse_number(text)
    if number is None:
        return None
    if "%" in text or number > 1.0:
        return number / 100.0
    return number


def format_number(value: Optional[float]) -> str:
    if value is None:
        return "-"
    if abs(value - round(value)) < 1e-9:
        return str(int(round(value)))
    return f"{value:.2f}"


def format_percent(value: Optional[float]) -> str:
    if value is None:
        return "-"
    return f"{value * 100:.2f}%"


def format_currency(value: Optional[float]) -> str:
    if value is None:
        return "-"
    return f"{value:.2f}"


def parse_price_candidates(price_text: str) -> List[float]:
    values = []
    for token in re.findall(r"\d+(?:\.\d+)?", price_text.replace(",", "")):
        try:
            values.append(float(token))
        except ValueError:
            continue
    return values[:2]


def percentile(values: Sequence[float], q: float) -> Optional[float]:
    if not values:
        return None
    ordered = sorted(values)
    if len(ordered) == 1:
        return ordered[0]
    position = (len(ordered) - 1) * (q / 100.0)
    lower = int(position)
    upper = min(lower + 1, len(ordered) - 1)
    if lower == upper:
        return ordered[lower]
    weight = position - lower
    return ordered[lower] * (1 - weight) + ordered[upper] * weight


def positive_numbers(values: Iterable[Optional[float]]) -> List[float]:
    return [float(value) for value in values if value is not None and float(value) > 0]


def mean_or_none(values: Sequence[float]) -> Optional[float]:
    return statistics.mean(values) if values else None


def strip_tags(value: str) -> str:
    text = re.sub(r"<[^>]+>", " ", value or "")
    return re.sub(r"\s+", " ", html_unescape(text)).strip()


def html_unescape(value: str) -> str:
    return unescape(value or "")
