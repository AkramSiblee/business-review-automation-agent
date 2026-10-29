import pickle
m = pickle.load(open("data/outputs/_model.pkl", "rb"))
pl, bs = m["consol_pl"], m["consol_bs"]
print(f'{"Period":8} {"Revenue":>12} {"HWRev":>10} {"GrossProfit":>12} {"GM%":>6} {"EBIT":>10} {"NetInc":>10} {"Cash":>12} {"BudgetRev":>12} {"RevVar%":>8}')
for p in m["PERIODS"]:
    c = pl[p]
    b = bs[p]
    bud = m["bud_revenue"].get(p, 0)
    var = (c["revenue"] - bud) / bud * 100 if bud else float("nan")
    print(f'{p:8} {c["revenue"]:12,.0f} {c.get("revenue_hardware",0):10,.0f} {c["gross_profit"]:12,.0f} {c["gross_profit"]/c["revenue"]*100:5.1f}% {c["ebit"]:10,.0f} {c["net_income"]:10,.0f} {b["cash"]:12,.0f} {bud:12,.0f} {var:7.1f}%')

print()
print("one-time detail:")
for p, items in m["one_time_detail"].items():
    if items:
        print(p, items)

print()
print("customer top 5:", m["top_customers"])
print("restated note:", m["restated_note"])
print()
print("headcount total by period (all depts):")
for p in m["PERIODS"]:
    total = sum(d["active"] for d in m["hc_by_period_dept"][p].values())
    hires = sum(d["hires"] for d in m["hc_by_period_dept"][p].values())
    terms = sum(d["terms_vol"] + d["terms_inv"] for d in m["hc_by_period_dept"][p].values())
    print(p, "active:", total, "hires:", hires, "terms:", terms)

print()
print("revenue bridge by period:")
for p in m["PERIODS"]:
    print(p, m["rev_bridge"][p])
