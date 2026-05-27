"""Sample listed + external-audit companies and check EBITDA quality."""
import sys
import types
import random
import os

# --- Stub streamlit so app.py can be imported without UI ---
fake_st = types.ModuleType("streamlit")

class _Sec(dict):
    def get(self, k, default=None):
        return os.environ.get(k, default)

fake_st.secrets = _Sec()

def _noop(*a, **k):
    return None

def _identity_decorator(*dargs, **dkwargs):
    def wrap(fn):
        # Add .clear() to mimic st.cache_*
        fn.clear = lambda: None
        return fn
    return wrap

class _Sidebar:
    def __getattr__(self, name):
        return _noop

fake_st.sidebar = _Sidebar()
fake_st.cache_data = _identity_decorator
fake_st.cache_resource = _identity_decorator

class _Stop(Exception):
    pass

fake_st.stop = lambda: (_ for _ in ()).throw(_Stop())
for nm in ("set_page_config", "markdown", "divider", "info", "warning", "error",
           "dataframe", "title", "header", "subheader", "write", "caption",
           "segmented_control", "text_input", "selectbox", "button", "columns",
           "spinner", "progress", "empty", "container", "expander", "tabs",
           "session_state", "rerun", "experimental_rerun", "form", "form_submit_button",
           "metric", "plotly_chart", "table", "code", "json", "image"):
    setattr(fake_st, nm, _noop)
fake_st.session_state = {}

sys.modules["streamlit"] = fake_st

# Stub st_aggrid (optional in app)
sys.modules.setdefault("st_aggrid", types.ModuleType("st_aggrid"))

os.environ["DART_API_KEY"] = "c0aacbfba7404217704ef01f2bdce5467a353fce"

# Now import app — guard the st.stop() call
import importlib.util
spec = importlib.util.spec_from_file_location("dartapp", "/home/user/DART-data/app.py")
mod = importlib.util.module_from_spec(spec)
try:
    spec.loader.exec_module(mod)
except _Stop:
    print("st.stop hit unexpectedly", file=sys.stderr)
    raise
except Exception as e:
    # The app may fail at top-level UI rendering after dart init — but by then
    # the functions we need are defined. We allow partial init.
    print(f"[import-warning] {type(e).__name__}: {e}", file=sys.stderr)


# --- Use module functions ---
dart = mod.get_dart(os.environ["DART_API_KEY"])
corp_df = mod.download_corp_code_xml(os.environ["DART_API_KEY"])
print(f"corp_code rows: {len(corp_df)}")

# Listed: stock_code non-empty (6 digits)
listed_mask = corp_df["stock_code"].astype(str).str.strip().str.len() == 6
listed = corp_df[listed_mask].copy()
# External audit candidates: stock_code empty (could include private; we'll verify via company())
external = corp_df[~listed_mask].copy()
print(f"listed candidates: {len(listed)}, non-listed: {len(external)}")

random.seed(42)

YEARS = [2021, 2022, 2023]
FS_DIV = "CFS"

def run_one(corp_code, corp_name, corp_cls_hint):
    try:
        # Determine corp_cls
        if corp_cls_hint:
            corp_cls = corp_cls_hint
        else:
            info = mod.enrich_company_info(dart, corp_code) or {}
            corp_cls = info.get("corp_cls", "E")
        yd, ym = mod.collect_multi_year(dart, corp_code, corp_cls, YEARS, FS_DIV)
        metrics = mod.compute_yearly_metrics(yd, ym, YEARS)
        result = []
        for y in YEARS:
            op = metrics["op_income"][y]
            eb = metrics["ebitda"][y]
            result.append((y, op, eb, ym.get(y, {}).get("source")))
        return result, yd, ym
    except Exception as e:
        return f"ERR {type(e).__name__}: {e}", None, None


def sample_pool(df, n, want_cls):
    picks = []
    pool = df.sample(frac=1, random_state=42).to_dict("records")
    for row in pool:
        if len(picks) >= n:
            break
        cc = row["corp_code"]
        try:
            info = mod.enrich_company_info(dart, cc) or {}
        except Exception:
            continue
        cls = info.get("corp_cls", "")
        if want_cls == "listed" and cls in ("Y", "K", "N"):
            picks.append((cc, row["corp_name"], cls))
        elif want_cls == "external" and cls in ("E",):
            picks.append((cc, row["corp_name"], cls))
    return picks


print("== sampling listed ==")
listed_pick = sample_pool(listed, 10, "listed")
for cc, nm, cls in listed_pick:
    print("  ", cc, cls, nm)

print("== sampling external ==")
external_pick = sample_pool(external, 10, "external")
for cc, nm, cls in external_pick:
    print("  ", cc, cls, nm)

import json
issues = []

def evaluate(label, picks):
    print(f"\n### {label} ###")
    for cc, nm, cls in picks:
        res, yd, ym = run_one(cc, nm, cls)
        if isinstance(res, str):
            print(f"[{cls}] {nm} ({cc}): {res}")
            continue
        for (y, op, eb, src) in res:
            tag = ""
            if eb is None and op is not None:
                tag = "EBITDA=None"
            elif eb is not None and op is not None and eb == op:
                tag = "EBITDA==OP"
            if tag:
                # Look at D&A keys
                row = yd.get(y, {})
                da_keys = {k: row.get(k) for k in (
                    "_유형자산감가상각비", "_무형자산상각비",
                    "_사용권자산상각비", "_감가상각비_합산")}
                meta = ym.get(y, {})
                print(f"[{cls}] {nm} ({cc}) {y}: {tag} op={op} eb={eb} src={src} "
                      f"da_source={meta.get('da_source')} D&A={da_keys}")
                issues.append({
                    "cls": cls, "name": nm, "corp_code": cc, "year": y,
                    "tag": tag, "op": op, "eb": eb, "src": src,
                    "da_source": meta.get("da_source"),
                    "da": da_keys,
                    "report": meta.get("report_nm"),
                    "rcept_no": meta.get("rcept_no"),
                })

evaluate("LISTED", listed_pick)
evaluate("EXTERNAL", external_pick)

print("\n=== ISSUES ===")
print(json.dumps(issues, ensure_ascii=False, indent=2, default=str))
