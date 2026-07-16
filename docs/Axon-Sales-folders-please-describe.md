# Axon — Sales folders (please add descriptions)

Axon's **Save / Download email** feature always files order emails under the **SOP** folder:

```
T:\IF\Sales \ <country code> \ SOP \ <year> \ <client> \ <order> \ <category>
```

The **country code** comes from the sender's country, and **year / client / order** are read from
the email — those are self-explanatory. We only need your help with **two** things below.

*(The other folders under a country — COMPLAINT, SERVICE, NON-SOP, PROSPECTION, SUSPECTS,
CLIENTS GENERAL, FPZ, "verslagen bezoek & tel" — are **not used by Download**, so they don't need
describing. PRICE ENQUIRY is treated as part of SOP.)*

---

## 1. Country codes (already confirmed — no action needed)

| Code | Country |
|------|---------|
| AB   | Belgium |
| AD   | Germany |
| AF   | France |
| AN   | Netherlands |

---

## 2. Inside an order — what goes in each folder? (confirmed)

An email inside an order is filed by **type** (Order vs Quotation) and by **who it is with**
(MC / MI / MS):

| Folder | What it is for |
|--------|----------------|
| Order | Emails that are orders |
| Quotation | Quotes/offers (not orders); also contains MC / MI / MS |
| MC | Correspondence with the **customer** (the client) |
| MI | Correspondence **internal** — our own group (Axon Group, Noviso, Almeco, PCA …) |
| MS | Correspondence with **suppliers** |

Axon now uses these rules to pre-pick the right folder (e.g. an order from the customer →
`Order\MC`, a supplier's quote → `Quotation\MS`); you can always change it before saving.
