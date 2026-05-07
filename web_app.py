import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from datetime import datetime
from pathlib import Path
import sys
import openpyxl
from openpyxl.styles import Font

sys.path.insert(0, str(Path(__file__).parent))
from scrape_brands import (
    scrape_epicenter, scrape_eva, node_available,
    write_store_sheet, write_summary_sheet,
)

st.set_page_config(
    page_title="Brand Stock Checker",
    page_icon="🔍",
    layout="wide",
)

# ── Helpers ───────────────────────────────────────────────

def make_excel(results, brand, checked_at):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Summary")
    write_summary_sheet(ws, results, brand, checked_at)
    for store, products in results.items():
        ws = wb.create_sheet(store)
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


def to_df(products):
    rows = []
    for p in products:
        rows.append({
            "Product Name":   p.get("name", ""),
            "SKU":            p.get("sku", ""),
            "Regular Price":  to_float(p.get("price", "")),
            "On Discount":    "Yes" if p.get("on_discount") else "No",
            "Discount Price": to_float(p.get("discount_price")) if p.get("on_discount") else None,
            "In Stock":       "Yes" if p.get("in_stock") is True else ("No" if p.get("in_stock") is False else "?"),
            "URL":            p.get("url", ""),
        })
    return pd.DataFrame(rows)


# ── Sidebar ───────────────────────────────────────────────

with st.sidebar:
    st.title("🔍 Brand Stock Checker")
    st.caption("Check product availability and prices across online stores")
    st.divider()

    st.subheader("1. Select stores")
    use_epicenter = st.checkbox("Epicenter (epicentrk.ua)", value=True)
    use_eva       = st.checkbox("Eva (eva.ua)",             value=True)

    st.divider()
    st.subheader("2. Enter brand name")
    brand = st.text_input(
        label="brand_input",
        value="Paclan",
        placeholder="e.g. Paclan",
        label_visibility="collapsed",
    )

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
        "👈 Select stores and enter a brand name in the sidebar, then click **Search**.\n\n"
        "The tool will scan all pages for that brand on each store and show a live preview "
        "with stock status, prices, and discounts. You can download the results as Excel."
    )
    st.stop()

# Validate inputs
if not brand.strip():
    st.error("Please enter a brand name.")
    st.stop()

selected = []
if use_epicenter: selected.append("Epicenter")
if use_eva:       selected.append("Eva")
if not selected:
    st.error("Please select at least one store.")
    st.stop()

# ── Scraping ──────────────────────────────────────────────

has_node   = node_available()
session    = requests.Session()
results    = {}
checked_at = datetime.now().strftime("%d.%m.%Y %H:%M")

for store in selected:
    with st.status(f"Scraping {store}...", expanded=True) as s:
        if store == "Epicenter":
            products = scrape_epicenter(
                brand.strip(), session, has_node, log_fn=st.write
            )
        else:
            products = scrape_eva(
                brand.strip(), session, log_fn=st.write
            )

        results[store] = products
        n      = len(products)
        in_n   = sum(1 for p in products if p["in_stock"] is True)
        disc_n = sum(1 for p in products if p.get("on_discount"))

        if n:
            s.update(
                label=f"✅ {store}: {n} products — {in_n} in stock, {disc_n} on discount",
                state="complete",
                expanded=False,
            )
        else:
            s.update(
                label=f"⚠️ {store}: no products found for '{brand.strip()}'",
                state="error",
                expanded=True,
            )

# ── Check for results ─────────────────────────────────────

total = sum(len(v) for v in results.values())
if total == 0:
    st.warning(f"No products found for **{brand.strip()}** in any selected store.")
    st.stop()

# ── Metrics ───────────────────────────────────────────────

st.divider()
total_in   = sum(sum(1 for p in v if p["in_stock"] is True)    for v in results.values())
total_out  = sum(sum(1 for p in v if p["in_stock"] is False)   for v in results.values())
total_disc = sum(sum(1 for p in v if p.get("on_discount"))     for v in results.values())

c1, c2, c3, c4 = st.columns(4)
c1.metric("Total Products", total)
c2.metric("In Stock",       total_in)
c3.metric("Out of Stock",   total_out)
c4.metric("On Discount",    total_disc)

# ── Preview table ─────────────────────────────────────────

st.divider()

stores_with_data = [s for s in results if results[s]]
tab_labels = stores_with_data + ["📋 Summary"]
tabs = st.tabs(tab_labels)

for i, store in enumerate(stores_with_data):
    with tabs[i]:
        df = to_df(results[store])

        # Filter controls
        col_search, col_stock, col_disc = st.columns([3, 1, 1])
        with col_search:
            search_text = st.text_input(
                "Filter by name",
                key=f"search_{store}",
                placeholder="Type to filter...",
                label_visibility="collapsed",
            )
        with col_stock:
            stock_filter = st.selectbox(
                "In Stock",
                ["All", "Yes", "No"],
                key=f"stock_{store}",
            )
        with col_disc:
            disc_filter = st.selectbox(
                "On Discount",
                ["All", "Yes", "No"],
                key=f"disc_{store}",
            )

        filtered = df.copy()
        if search_text:
            filtered = filtered[
                filtered["Product Name"].str.contains(search_text, case=False, na=False)
            ]
        if stock_filter != "All":
            filtered = filtered[filtered["In Stock"] == stock_filter]
        if disc_filter != "All":
            filtered = filtered[filtered["On Discount"] == disc_filter]

        st.caption(f"Showing {len(filtered)} of {len(df)} products")

        st.dataframe(
            filtered,
            use_container_width=True,
            hide_index=True,
            height=500,
            column_config={
                "Regular Price":  st.column_config.NumberColumn(
                    "Regular Price (UAH)", format="%.2f"
                ),
                "Discount Price": st.column_config.NumberColumn(
                    "Discount Price (UAH)", format="%.2f"
                ),
                "URL": st.column_config.LinkColumn(
                    "Product URL", display_text="Open ↗"
                ),
            },
        )

with tabs[-1]:
    rows = []
    for store, prods in results.items():
        in_n   = sum(1 for p in prods if p["in_stock"] is True)
        out_n  = sum(1 for p in prods if p["in_stock"] is False)
        disc_n = sum(1 for p in prods if p.get("on_discount"))
        rows.append({
            "Store":         store,
            "Total Products": len(prods),
            "In Stock":      in_n,
            "Out of Stock":  out_n,
            "On Discount":   disc_n,
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ── Download ──────────────────────────────────────────────

st.divider()
ts    = datetime.now().strftime("%Y%m%d_%H%M")
fname = f"{brand.strip().replace(' ', '_')}_{ts}.xlsx"

st.download_button(
    label="📥  Download Excel",
    data=make_excel(results, brand.strip(), checked_at),
    file_name=fname,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
)
