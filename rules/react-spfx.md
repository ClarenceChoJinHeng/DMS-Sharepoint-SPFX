# React / SPFx

- Every hook goes above the first early `return`. `#code/hooks-above-early-return`
- No backticks inside a CSS template literal (they end it).
- `position: fixed` breaks inside `container-type` or transforms; popovers get clipped by scroll boxes.
- `Promise.allSettled` and `.finally()` are not available (ES target).
- Empty ≠ unknown: a failed read is `undefined`, never `[]`. Gates on forms fail open; destructive/security decisions fail closed.
