# Axon — Sales folder structure (please add descriptions)

To make Axon's **Save / Download email** feature file emails into the correct folder automatically,
we need to understand what each folder in `T:\IF\Sales` is used for.

Below is the **actual folder tree**. The **year / client / order** levels are self-explanatory
(one folder per year, per client, per order), so those don't need describing. Please fill in the
**"What it is for"** column in sections 1–3 for the folders that carry a meaning.

---

## The actual tree

```
T:\IF\Sales\
├─ AB\   (Belgium – confirmed)
│   ├─ CLIENTS GENERAL\
│   ├─ COMPLAINT\
│   ├─ NON-SOP\
│   ├─ PROSPECTION\
│   ├─ SERVICE\
│   ├─ SOP\            <-- orders (see "Inside SOP" below)
│   └─ SUSPECTS\
├─ AD\   (country = ?)
│   ├─ CLIENTS GENERAL\
│   ├─ COMPLAINT\
│   ├─ NON-SOP\
│   ├─ PRICE ENQUIRY\
│   ├─ PROSPECTION\
│   ├─ SOP\
│   └─ SUSPECTS\
├─ AF\   (country = ?)
│   ├─ CLIENTS GENERAL\
│   ├─ COMPLAINT\
│   ├─ NON-SOP\
│   ├─ PRICE ENQUIRY\
│   ├─ PROSPECTION\
│   ├─ SOP\
│   └─ SUSPECTS\
└─ AN\   (Netherlands – confirmed)
    ├─ CLIENTS GENERAL\
    ├─ COMPLAINT\
    ├─ FPZ\
    ├─ NON-SOP\
    ├─ PRICE ENQUIRY\
    ├─ PROSPECTION\
    ├─ SOP\
    ├─ SUSPECTS\
    └─ verslagen bezoek & tel\

Inside every SOP\ (orders):

  SOP\
  └─ <year>\                one per year        e.g. "2026 (14334-)", "2025 (14094-)"
     └─ <client>\           one per client      e.g. Voestalpine, Agristo, Actemium
        └─ <order>\         one per order       e.g. 14457_ADD low noise  (number_description)
           ├─ Documents\
           ├─ Order\        (may contain: MC, MI, MS, PO, SO, Internal, or a supplier name)
           ├─ Quotation\    (may contain: MC, MS, ...)
           ├─ MC\
           ├─ MI\
           └─ MS\
```

---

## 1. Country codes — which country is each?

| Code | Country |
|------|---------|
| AB   | Belgium (confirmed) |
| AD   |  |
| AF   |  |
| AN   | Netherlands (confirmed) |

---

## 2. Type folders — what is each used for? (Which one holds order / sales emails?)

| Folder | What it is for |
|--------|----------------|
| SOP |  |
| COMPLAINT |  |
| SERVICE  (AB only) |  |
| NON-SOP |  |
| PROSPECTION |  |
| SUSPECTS |  |
| CLIENTS GENERAL |  |
| PRICE ENQUIRY  (AD, AF, AN) |  |
| FPZ  (AN only) |  |
| verslagen bezoek & tel  (AN only) |  |

---

## 3. Inside an order — what goes in each?

| Folder | What it is for |
|--------|----------------|
| Documents |  |
| Order |  |
| Quotation |  |
| MC |  |
| MI |  |
| MS |  |
| PO   (seen inside Order) |  |
| SO   (seen inside Order) |  |
| Internal |  |

---

## Anything we missed?

| Folder / path | What it is for |
|---------------|----------------|
|  |  |
|  |  |

---

When you return this, Axon will use **your** definitions to pre-pick the right type folder and
category for each email — while you can always change it before saving.
