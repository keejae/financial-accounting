# Philip Hong — Accounting Learning Resources

A static website hosting course materials for **Introduction to Financial Accounting**, taught by Philip Hong, CPA, Ph.D. Deployed via GitHub Pages at [www.philiphong.org](https://www.philiphong.org).

Repository: <https://github.com/keejae/financial-accounting/>

---

## Three-level navigation

Students are never more than 3 clicks from any file:

1. **Home** (`index.html`) — welcome + current courses + latest updates
2. **Courses** (`courses.html`) — the course landing page with weekly cards
3. **Week materials** (`weeks/weekN.html`) — Lecture Notes + Homework cards linking to downloadable files

Additional page: **About** (biography + CV download / view online).

## Site tree

```
financial-accounting/
├── index.html              Home
├── courses.html            Course landing (Level 2)
├── about.html              About + CV download/view
├── CNAME                   GitHub Pages custom domain
├── .nojekyll               Disable Jekyll processing
├── Photo.png               Portrait shown on About page
├── CV_Philip_Keejae_Hong_April 2026.pdf   Linked from About page
├── css/style.css
├── js/main.js              Language toggle + mobile menu
├── images/
├── weeks/                  Level 3
│   ├── week1.html   … week4.html
│   └── exam.html
└── files/                  Downloadable content (per-language folders)
    ├── week1/english/, korean/
    ├── week2/english/, korean/
    ├── week3/english/, korean/
    ├── week4/english/, korean/
    └── exam/english/, korean/
```

---

## How to add course materials

Drop files into `files/<week>/<lang>/` using these exact filenames:

**Week folders (week1–week4):**
| Filename | Purpose |
|----------|---------|
| `lecture-notes.pdf` | PDF version (also opened by the "View" button in-browser) |
| `lecture-notes.docx` | Editable Word version |
| `slides.pptx` | PowerPoint slides |
| `practice-problems.pdf` | Weekly homework |
| `solutions.pdf` | Homework solutions (post when ready) |
| `recording.html` | Already provided — edit to embed the recording |

**Exam folder:**
| Filename | Purpose |
|----------|---------|
| `exam-review.pdf` | Review sheet with key concepts and formulas |
| `practice-problems.pdf` | Practice exam questions |
| `practice-solutions.pdf` | Solutions (post when ready) |
| `additional-materials.zip` | Formula sheet, sample exams, etc. |
| `recording.html` | Already provided |

**CV folder:**
Drop `files/cv/cv.pdf`. The CV page embeds it as an iframe and offers Download / View Online buttons.

### Embedding a lecture recording

Open `files/<week>/<lang>/recording.html` and follow the comment at the top:

- **YouTube:** paste the video ID into the `<iframe>` URL and remove the "notice" block.
- **Zoom:** paste the Zoom share URL into an `<iframe src="…">`.
- **MP4:** replace the `<div class="video-embed">` block with `<video controls src="recording.mp4"></video>` and place the file next to `recording.html`.

---

## Editing content

- **Professor's Notes** at the top of each week page: edit the `<p>` inside `<div class="prof-note">` in `weeks/weekN.html`.
- **Week topics** in the "Course Materials" section: edit `weeks/weekN.html` (title + eyebrow) *and* the matching card in `courses.html`.
- **Latest Updates** on the home page: edit the `<ul class="updates-list">` in `index.html`.
- **Language toggle** remembers the visitor's last choice via `localStorage`.

## Adding Week 5 later

1. Copy `weeks/week4.html` to `weeks/week5.html`. Update title, eyebrow, topic, professor's note, prev/next nav, and file paths (`/week4/` → `/week5/`).
2. Add a new `<a class="week-card">` on `courses.html`.
3. Update the "Next" link on `week4.html` to point to `week5.html`.
4. Create `files/week5/english/` and `files/week5/korean/` folders.

## Adding a second course later

1. Duplicate `courses.html` to e.g. `intermediate-accounting.html` and update its cards/links.
2. Rename the current `courses.html` to `financial-accounting.html`.
3. Create a new `courses.html` that lists both courses (index page).
4. Update the nav in every page: `courses.html` → still points to the courses index.

---

## Deploying to GitHub Pages

1. Push this folder to <https://github.com/keejae/financial-accounting/>.
2. **Settings → Pages** → Source: **Deploy from a branch**, branch **main**, folder `/ (root)`.
3. Save. GitHub builds the site.

### Custom domain (www.philiphong.org)

The `CNAME` file at repo root already contains `www.philiphong.org`.

At your domain registrar, create DNS records:

| Type  | Host   | Value                        |
|-------|--------|------------------------------|
| CNAME | www    | keejae.github.io             |
| A     | @      | 185.199.108.153              |
| A     | @      | 185.199.109.153              |
| A     | @      | 185.199.110.153              |
| A     | @      | 185.199.111.153              |

The `A` records at the apex (`@`) redirect `philiphong.org` → `www.philiphong.org`. DNS propagation can take up to 24 hours.

In **Settings → Pages** on GitHub, confirm the custom domain shows `www.philiphong.org` and enable **Enforce HTTPS** once the certificate is provisioned.

---

## Local preview

Open `index.html` directly in a browser — everything works over `file://`. Resize the window below ~768px to see the mobile hamburger menu.

## Colors & styling

Colors, spacing, and typography are defined as CSS variables at the top of `css/style.css`. The primary accent color is CMU Tartan Red (`#C41230`); all corners are sharp (no border-radius).

&copy; Professor Philip Hong
