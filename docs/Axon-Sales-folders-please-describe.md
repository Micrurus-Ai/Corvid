# Axon — Sales folder structure (please add descriptions)

To make Axon's **Save / Download email** feature file emails into the correct folder
automatically, we need to understand what each folder in `T:\IF\Sales` is used for.

Below is the structure. **Please fill in the "What it is for" column** for the folders in
sections 1–3. The year, client and order folders are self-explanatory, so those don't need
describing — only the ones marked below.

---

## The structure (the repeating pattern)

```
T:\IF\Sales\
  <country code>\          e.g. AB, AD, AF, AN
    <type>\                e.g. SOP, COMPLAINT, SERVICE ...      <-- section 2
      <year>\              e.g. 2026 (14334-)                    (self-explanatory)
        <client>\          e.g. Voestalpine, Agristo            (self-explanatory)
          <order>\         e.g. 14457_ADD low noise             (number_description)
            <category>\    e.g. Documents, Order, Quotation      <-- section 3
              <sub>\       e.g. MC, MI, MS                       <-- section 3
```

---

## 1. Country codes — which country is each?

| Code | Country              |
|------|----------------------|
| AB   | Belgium (confirmed)  |
| AD   |                      |
| AF   |                      |
| AN   | Netherlands (confirmed) |

---

## 2. Type folders — what is each used for?

These sit directly under each country code. **Which one holds order / sales emails?**
(Some appear only in certain countries, noted in brackets.)

| Folder                     | What it is for |
|----------------------------|----------------|
| SOP                        |                |
| COMPLAINT                  |                |
| SERVICE                    |                |
| NON-SOP                    |                |
| PROSPECTION                |                |
| SUSPECTS                   |                |
| CLIENTS GENERAL            |                |
| PRICE ENQUIRY  (AD, AF, AN) |               |
| FPZ  (AN only)             |                |
| verslagen bezoek & tel  (AN only) |         |

---

## 3. Inside an order — what goes in each?

These sit inside an order folder (e.g. inside `14457_ADD low noise`).

| Folder     | What it is for |
|------------|----------------|
| Documents  |                |
| Order      |                |
| Quotation  |                |
| MC         |                |
| MI         |                |
| MS         |                |

---

## Anything we missed?

If there are other important folders, or the meaning changes by country, please add them here:

| Folder / path | What it is for |
|---------------|----------------|
|               |                |
|               |                |

---

With these descriptions, Axon will know, for each incoming email, **which type folder and which
category** it belongs in — so Save/Download files it in the right place automatically.
