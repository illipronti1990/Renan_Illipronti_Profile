# AGENTS.md

## Cursor Cloud specific instructions

This repository is a **GitHub profile README** (personal "About Me" page). On the
`main` branch it contains a single tracked file, `README.md`. There is:

- No application code, backend, or frontend
- No package manager, lockfile, or dependency manifest
- No build, lint, test, or service to run

Because of this, there is **nothing to install** and the update script is a no-op.

### Previewing the README (the only meaningful "dev" activity)

`README.md` is Markdown that GitHub renders automatically on the profile page. It
uses external image services (shields.io, capsule-render, github-readme-stats,
streak-stats, komarev) referenced purely as image URLs.

To preview locally the way GitHub renders it:

1. Render to GitHub-flavored HTML: `gh api --method POST /markdown/raw -H "Content-Type: text/plain" --input README.md > body.html`
2. Wrap `body.html` in a minimal HTML page and serve it, e.g. `python3 -m http.server 8899`
3. Open `http://localhost:8899/` in a browser.

Notes:
- The `github-readme-stats` and `streak-stats` widgets query the GitHub API
  unauthenticated and are rate-limited from arbitrary IPs, so they may show a
  "Failed to retrieve contributions" placeholder locally. This is an external
  service limitation, not a repo problem — they render fine on the live profile.
- Do not add build tooling or dependencies to this repo just to preview it.
