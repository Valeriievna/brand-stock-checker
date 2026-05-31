import streamlit as st
import pandas as pd
import requests
import random
from io import BytesIO
from datetime import datetime
from pathlib import Path
import sys
import openpyxl
from openpyxl.styles import Font

sys.path.insert(0, str(Path(__file__).parent))
from scrape_brands import (
    scrape_epicenter, scrape_eva, scrape_organic, node_available,
    write_store_sheet, write_summary_sheet, check_data_quality,
)

st.set_page_config(
    page_title="Brand Stock Checker",
    page_icon="🔍",
    layout="wide",
)

PRESET_BRANDS = [
    "Paclan", "Vileda", "Domi", "Фрекен Бок", "FINO",
    "Stella", "Spontex", "PRO SERVIS",
    "Добра Господарка", "York", "Помічниця",
]

STORE_BRANDS = {
    "Epicenter":      PRESET_BRANDS,
    "Eva":            PRESET_BRANDS,
    "Organic Market": ["Paclan"],
}

# ── Helpers ───────────────────────────────────────────────

def make_excel(all_results, checked_at):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Summary sheet
    ws_sum = wb.create_sheet("Summary")
    hdr_font = Font(bold=True, color="FFFFFF")
    from openpyxl.styles import PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    hdr_fill = PatternFill("solid", fgColor="2E75B6")
    center   = Alignment(horizontal="center", vertical="center")

    headers = ["Brand", "Store", "Total", "In Stock", "Out of Stock", "Expected", "On Discount"]
    widths  = [20, 22, 10, 12, 14, 12, 13]
    ws_sum.merge_cells(f"A1:{get_column_letter(len(headers))}1")
    t = ws_sum["A1"]
    t.value     = f"Summary  ·  {checked_at}"
    t.font      = Font(bold=True, size=13, color="1F4E79")
    t.alignment = center
    ws_sum.row_dimensions[1].height = 26
    for ci, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws_sum.cell(row=2, column=ci, value=h)
        c.font = hdr_font; c.fill = hdr_fill; c.alignment = center
        ws_sum.column_dimensions[get_column_letter(ci)].width = w

    ri = 3
    for brand, store_results in all_results.items():
        for store, products in store_results.items():
            in_n  = sum(1 for p in products if p["in_stock"] is True)
            out_n = sum(1 for p in products if p["in_stock"] is False)
            exp_n = sum(1 for p in products if p["in_stock"] == "expected")
            disc_n = sum(1 for p in products if p.get("on_discount"))
            for ci, val in enumerate([brand, store, len(products), in_n, out_n, exp_n, disc_n], 1):
                ws_sum.cell(row=ri, column=ci, value=val).alignment = center
            ri += 1

    # Per-brand+store sheets
    for brand, store_results in all_results.items():
        for store, products in store_results.items():
            sheet_name = f"{brand[:15]} - {store[:14]}"
            ws = wb.create_sheet(sheet_name)
            if products:
                write_store_sheet(ws, products, store, brand, checked_at)
            else:
                ws["A1"] = f'No products found for "{brand}" on {store}.'
                ws["A1"].font = Font(bold=True, color="FF0000")

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def to_float(val):
    try:
        return float(str(val).replace(" ", "").replace(",", "."))
    except (ValueError, AttributeError):
        return None


def stock_label(p):
    v = p.get("in_stock")
    if v is True:      return "Yes"
    if v is False:     return "No"
    if v == "expected": return "Expected"
    return "?"


def to_df(products):
    rows = []
    for p in products:
        rows.append({
            "Product Name":   p.get("name", ""),
            "Category":       p.get("category", ""),
            "SKU":            p.get("sku", ""),
            "Regular Price":  to_float(p.get("price", "")),
            "On Discount":    "Yes" if p.get("on_discount") else "No",
            "Discount Price": to_float(p.get("discount_price")) if p.get("on_discount") else None,
            "In Stock":       stock_label(p),
            "Seller":         p.get("seller", ""),
            "URL":            p.get("url", ""),
        })
    return pd.DataFrame(rows)


DF_COL_CONFIG = {
    "Regular Price":  st.column_config.NumberColumn("Regular Price (UAH)", format="%.2f"),
    "Discount Price": st.column_config.NumberColumn("Discount Price (UAH)", format="%.2f"),
    "URL":            st.column_config.LinkColumn("Product URL", display_text="Open ↗"),
}


# ── Sidebar ───────────────────────────────────────────────

with st.sidebar:
    st.title("🔍 Brand Stock Checker")
    st.caption("Check product availability and prices across online stores")
    st.divider()

    st.subheader("1. Select stores")
    use_epicenter = st.checkbox("Epicenter (epicentrk.ua)",               value=True)
    use_eva       = st.checkbox("Eva (eva.ua)",                           value=True)
    use_organic   = st.checkbox("Organic Market (organic-market.com.ua)", value=True)

    selected_stores_sidebar = []
    if use_epicenter: selected_stores_sidebar.append("Epicenter")
    if use_eva:       selected_stores_sidebar.append("Eva")
    if use_organic:   selected_stores_sidebar.append("Organic Market")

    st.divider()
    st.subheader("2. Select brands")

    # Show only brands available on at least one selected store
    available_brands = [
        b for b in PRESET_BRANDS
        if any(b in STORE_BRANDS.get(s, []) for s in selected_stores_sidebar)
    ]

    selected_brands = []
    for b in available_brands:
        if st.checkbox(b, value=(b == "Paclan"), key=f"brand_{b}"):
            selected_brands.append(b)

    custom_brand = st.text_input(
        label="custom_brand",
        placeholder="Add custom brand...",
        label_visibility="collapsed",
    )
    if custom_brand.strip() and custom_brand.strip() not in selected_brands:
        selected_brands = list(selected_brands) + [custom_brand.strip()]

    st.divider()
    go = st.button("🔎  Search", type="primary", use_container_width=True)

    st.divider()
    if not node_available():
        st.warning(
            "⚠️ Node.js not found\n\n"
            "Epicenter scraping requires Node.js.\n"
            "Install from [nodejs.org](https://nodejs.org/)"
        )

# ── Main ──────────────────────────────────────────────────

st.title("Brand Stock Checker")

if not go:
    st.info(
        "👈 Select stores and brands in the sidebar, then click **Search**.\n\n"
        "You can select multiple brands at once. The tool will scan each brand "
        "across all selected stores and show stock status, prices, and discounts."
    )
    st.stop()

# Validate
if not selected_brands:
    st.error("Please select at least one brand.")
    st.stop()

selected_stores = selected_stores_sidebar
if not selected_stores:
    st.error("Please select at least one store.")
    st.stop()

# ── Scraping ──────────────────────────────────────────────

has_node   = node_available()
session    = requests.Session()
checked_at = datetime.now().strftime("%d.%m.%Y %H:%M")

all_results = {}
all_metas   = {}

for brand in selected_brands:
    all_results[brand] = {}
    all_metas[brand]   = {}
    for store in selected_stores:
        with st.status(f"Scraping {store} — {brand}...", expanded=True) as s:
            meta = {}
            if store == "Epicenter":
                products = scrape_epicenter(brand, session, has_node, log_fn=st.write, meta=meta)
            elif store == "Eva":
                products = scrape_eva(brand, session, log_fn=st.write, meta=meta)
            else:
                products = scrape_organic(brand, session, log_fn=st.write, meta=meta)

            all_results[brand][store] = products
            all_metas[brand][store]   = meta
            n      = len(products)
            in_n   = sum(1 for p in products if p["in_stock"] is True)
            disc_n = sum(1 for p in products if p.get("on_discount"))

            if n:
                s.update(
                    label=f"✅ {brand} / {store}: {n} products — {in_n} in stock, {disc_n} on discount",
                    state="complete", expanded=False,
                )
            else:
                s.update(
                    label=f"⚠️ {brand} / {store}: no products found",
                    state="error", expanded=True,
                )

# ── Check for results ─────────────────────────────────────

total_all = sum(
    len(prods)
    for store_res in all_results.values()
    for prods in store_res.values()
)
if total_all == 0:
    st.warning("No products found for any selected brand in any selected store.")
    st.stop()

# ── Metrics ───────────────────────────────────────────────

st.divider()
all_products_flat = [p for sr in all_results.values() for prods in sr.values() for p in prods]

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total Products", len(all_products_flat))
c2.metric("In Stock",       sum(1 for p in all_products_flat if p["in_stock"] is True))
c3.metric("Out of Stock",   sum(1 for p in all_products_flat if p["in_stock"] is False))
c4.metric("Expected",       sum(1 for p in all_products_flat if p["in_stock"] == "expected"))
c5.metric("On Discount",    sum(1 for p in all_products_flat if p.get("on_discount")))

# ── Data Quality ──────────────────────────────────────────

st.divider()
for brand, store_results in all_results.items():
    for store, products in store_results.items():
        if not products:
            continue
        meta       = all_metas.get(brand, {}).get(store, {})
        site_total = meta.get("site_total")
        scraped    = len(products)
        warnings   = check_data_quality(products)
        count_ok   = site_total is None or scraped >= site_total
        data_ok    = len(warnings) == 0

        if site_total or warnings:
            with st.expander(
                f"{'✅' if count_ok and data_ok else '⚠️'} Data quality — {brand} / {store}",
                expanded=not (count_ok and data_ok),
            ):
                if site_total:
                    if count_ok:
                        st.success(f"Count check: found {scraped}, site reports {site_total} ✅")
                    else:
                        st.warning(f"Count check: found {scraped}, site reports {site_total} — may be incomplete ⚠️")
                if warnings:
                    st.warning(f"{len(warnings)} suspicious product(s):")
                    for msg, url in warnings:
                        st.caption(f"• {msg} — [Open ↗]({url})" if url else f"• {msg}")

                sample = random.sample(products, min(5, len(products)))
                st.markdown("**Spot-check sample:**")
                spot_rows = [{"Product Name": p.get("name",""), "Regular Price": to_float(p.get("price","")),
                               "Discount Price": to_float(p.get("discount_price")) if p.get("on_discount") else None,
                               "In Stock": stock_label(p), "Seller": p.get("seller",""), "URL": p.get("url","")}
                              for p in sample]
                st.dataframe(pd.DataFrame(spot_rows), use_container_width=True, hide_index=True,
                             column_config={"Regular Price": st.column_config.NumberColumn("Regular Price (UAH)", format="%.2f"),
                                            "Discount Price": st.column_config.NumberColumn("Discount Price (UAH)", format="%.2f"),
                                            "URL": st.column_config.LinkColumn("Link", display_text="Open ↗")})

# ── Preview tabs ──────────────────────────────────────────

st.divider()

brands_with_data = [b for b in all_results if any(all_results[b].values())]
tab_labels = brands_with_data + ["📋 Summary"]
tabs = st.tabs(tab_labels)

for ti, brand in enumerate(brands_with_data):
    with tabs[ti]:
        store_results = all_results[brand]
        stores_with_data = [s for s in store_results if store_results[s]]

        if not stores_with_data:
            st.warning(f"No products found for {brand}.")
            continue

        store_tabs = st.tabs(stores_with_data)
        for si, store in enumerate(stores_with_data):
            with store_tabs[si]:
                df = to_df(store_results[store])

                col_search, col_stock, col_disc = st.columns([3, 1, 1])
                with col_search:
                    search_text = st.text_input("Filter", key=f"s_{brand}_{store}",
                                                placeholder="Filter by name...", label_visibility="collapsed")
                with col_stock:
                    stock_filter = st.selectbox("Stock", ["All","Yes","No","Expected"],
                                                key=f"st_{brand}_{store}")
                with col_disc:
                    disc_filter = st.selectbox("Discount", ["All","Yes","No"],
                                               key=f"d_{brand}_{store}")

                filtered = df.copy()
                if search_text:
                    filtered = filtered[filtered["Product Name"].str.contains(search_text, case=False, na=False)]
                if stock_filter != "All":
                    filtered = filtered[filtered["In Stock"] == stock_filter]
                if disc_filter != "All":
                    filtered = filtered[filtered["On Discount"] == disc_filter]

                st.caption(f"Showing {len(filtered)} of {len(df)} products")
                st.dataframe(filtered, use_container_width=True, hide_index=True,
                             height=500, column_config=DF_COL_CONFIG)

with tabs[-1]:
    rows = []
    for brand, store_results in all_results.items():
        for store, products in store_results.items():
            rows.append({
                "Brand":          brand,
                "Store":          store,
                "Total Products": len(products),
                "In Stock":       sum(1 for p in products if p["in_stock"] is True),
                "Out of Stock":   sum(1 for p in products if p["in_stock"] is False),
                "Expected":       sum(1 for p in products if p["in_stock"] == "expected"),
                "On Discount":    sum(1 for p in products if p.get("on_discount")),
            })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ── Download ──────────────────────────────────────────────

st.divider()
ts     = datetime.now().strftime("%Y%m%d_%H%M")
brands_str = "_".join(b.replace(" ", "") for b in selected_brands)[:40]
fname  = f"{brands_str}_{ts}.xlsx"

st.download_button(
    label="📥  Download Excel",
    data=make_excel(all_results, checked_at),
    file_name=fname,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
)
