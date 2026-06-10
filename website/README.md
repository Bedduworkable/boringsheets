# BoringSheets — marketing site

A fast, static, SEO-optimized landing site for BoringSheets, built with plain
HTML/CSS and GSAP for the animations (nav overlay, SplitText char-flip, scroll
line-reveal). No build step — deploy the folder as-is.

## Files
- `index.html` — landing page (hero, features, engine, download, FAQ) + full SEO + JSON-LD
- `privacy.html` / `terms.html` — legal pages
- `styles.css` — design system (dark + "shockingly green")
- `main.js` — GSAP interactions
- `robots.txt`, `sitemap.xml`, `site.webmanifest` — SEO/PWA
- `assets/favicon.svg`, `assets/og.svg` — icon + social share image

## Run locally
```bash
cd website
python3 -m http.server 8080
# open http://localhost:8080
```

## Deploy
Drop the `website/` folder on any static host: **Netlify, Vercel, Cloudflare
Pages, GitHub Pages, or S3**. No server code required.

## Before you go live — replace these placeholders
1. **Domain.** Every absolute URL uses `https://boringsheets.app`. Find & replace
   it with your real domain across `index.html`, `privacy.html`, `terms.html`,
   `robots.txt`, and `sitemap.xml`.
2. **Download links.** The "Download for Mac" buttons (`data-dl="mac"` and the
   one in `#download`) point to `#`. Point them at your real `.dmg` URL.
3. **Social links.** GitHub / X links in the nav and footer are `#` placeholders.
4. **OG image.** `assets/og.svg` works, but some platforms (Facebook, some
   Twitter/X cards) prefer a raster image. Export a **1200×630 PNG** as
   `assets/og.png` and update the `og:image` / `twitter:image` tags.
5. **Analytics (optional).** None is included (privacy-first). If you add any,
   update `privacy.html` accordingly.

## SEO included
- Title, meta description, keywords, canonical, theme-color, robots
- Open Graph + Twitter Card tags
- JSON-LD: `Organization`, `WebSite`, `SoftwareApplication`, `FAQPage`
- Semantic landmarks (`header`, `nav`, `main`, `section`, `footer`), `alt`/`aria`
- `sitemap.xml` + `robots.txt`
- Fast: system fonts, deferred scripts, no render-blocking CSS beyond one file
