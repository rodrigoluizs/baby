# Baby Pool Backend — Setup Guide

The betting pool stores bets in a Google Sheet through a Google Apps Script
Web App. No servers, no cost. One-time setup, about 5 minutes.

## 1. Create the spreadsheet

1. Go to [sheets.new](https://sheets.new) and name it e.g. **Baby Pool**.
2. Menu **Extensions → Apps Script**.
3. Delete the sample code and paste the full content of [`Code.gs`](Code.gs).
4. Save (Cmd+S).

## 2. Deploy as Web App

1. Click **Deploy → New deployment**.
2. Gear icon → type **Web app**.
3. Settings:
   - Description: `baby pool v1`
   - Execute as: **Me**
   - Who has access: **Anyone**
4. Click **Deploy**, authorize the permissions, and copy the **Web app URL**
   (looks like `https://script.google.com/macros/s/AKfy.../exec`).

## 3. Wire it into the site

In `index.html`, find:

```js
const POOL_URL = "";
```

and paste the Web app URL between the quotes. Commit and deploy. While the
URL is empty the whole pool feature stays hidden on the site.

The `bets` and `actual` sheets are created automatically on the first request.
Each bet is one row in `bets` — that sheet is your admin view.

## 4. When the baby arrives 👶

Fill **row 2** of the `actual` sheet:

| column         | what to enter                              | example      |
| -------------- | ------------------------------------------ | ------------ |
| `sex`          | `boy` or `girl` (dropdown)                 | `girl`       |
| `birth_date`   | date of birth                              | `07/09/2026` |
| `birth_time`   | time of birth (24h)                        | `14:35`      |
| `weight_grams` | weight in grams                            | `3450`       |
| `length_cm`    | length in cm                               | `51`         |
| `hair`         | `bald`, `some` or `lots` (dropdown)        | `some`       |
| `eyes`         | `blue`, `brown`, `green` or `black` (dropdown) | `brown`  |

Gestational week, weekday and day/night are derived automatically from the
date and time. As soon as every column is filled, the site switches from the
live results view to the final leaderboard — nothing else to do.

## Updating the script later

Edit the code in Apps Script, then **Deploy → Manage deployments → edit
(pencil) → Version: New version → Deploy**. Reusing the same deployment keeps
the URL stable; a *new* deployment would generate a different URL and require
updating `index.html`.
