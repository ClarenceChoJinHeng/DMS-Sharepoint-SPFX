---
name: sdg-mobile-css
description: Use when adding or changing layout or styling on any web part, or when the user reports something looks wrong on mobile or in the SharePoint mobile preview.
---
# Mobile-first CSS

1. **Write the phone layout first** (single column, full width, no side padding), then widen with `@media (min-width: ...)` or `@container (min-width: ...)`.
2. **Use `className` + a CSS block** for new styles. Inline style objects cannot hold media queries.
   - In a CSS template literal, **never use backticks**, not even in comments.
3. **Media vs container query:**
   - `@container` follows the web part's own width (the SharePoint section). Prefer it.
   - But `container-type` breaks any `position: fixed` child (dialogs, toasts). Put the container on an inner wrapper that ends **before** the dialogs, or use `@media`.
4. **Don't change desktop by accident:**
   - `max-width` caps are fine as they are.
   - Moving padding across a `max-width` element changes desktop width. Measure before and after.
   - Use `overflow-wrap: break-word`, not `anywhere`.
5. **Scroll boxes clip popovers.** Check for absolutely positioned children before adding `overflow` or `max-height`.
6. Test at 360px wide and at desktop width. Say which one was actually checked.
