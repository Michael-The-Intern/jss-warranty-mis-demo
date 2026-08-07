# JSS Warranty MIS Public Package B1+B2 Packaging Report v1.0

## Status

**PUBLIC B1+B2 ARTIFACT ACCEPTED FOR PUBLIC DEMO BRANCH TESTING**

## Inputs

### Public baseline

- File: `Download corrected public v1.3 JSX.jsx`
- SHA-256: `ecbaf9e02343d9b477d6b80610975b01fa4e45cd0093c6b1899d93889ef1b23c`
- Bytes: `373,408`
- Logical lines: `6,066`

### Accepted private Package B2 source

- File: `Download Package B2 Pass 6.jsx`
- SHA-256: `384e649f90a4e48e64fc9b12a27aa0ab769a5342a4711e8d19a14ed4e8c37666`
- Bytes: `438,682`
- Logical lines: `7,239`

## Output

- File: `App_Public_Package_B1_B2_Profile_Preview_v1.0.jsx`
- SHA-256: `a81e453ff139ff611c839cc7e4e7642617b8e8dde888fcf4a6b11596cfae0abd`
- Bytes: `438,874`
- Logical lines: `7,242`

## Packaging method

Only the complete Phase 2 Operations / Import Center block was transplanted from the accepted private Package B2 source into the verified public v1.3 baseline.

Verification:

- Public source before Phase 2: byte-identical.
- Public source after Phase 2: byte-identical.
- Transplanted Phase 2 block: byte-identical to accepted private Pass 6.
- No complete-private-App replacement occurred.

## Static and privacy checks

- TypeScript JSX syntax diagnostics: 0 errors.
- Single `WarrantyMISImportCenter` definition: PASS.
- Single top-level `App` router: PASS.
- Profile Preview transition present: PASS.
- Mapped Profile Preview present: PASS.
- Original Source Preview present: PASS.
- Validation Review visible but disabled: PASS.
- Phase 2 fetch calls: none.
- Phase 2 Axios calls: none.
- Phase 2 Supabase API calls: none.
- Phase 2 browser-storage API calls: none.
- Obvious repository tokens, service-role keys, and database URLs: none.
- Case # to Campaign Number dual-label contract: present.
- Debit number range remains batch context: present.
- Individual Debit records are not manufactured: confirmed.

## Build verification

An isolated Vite 8.2.1 production build passed:

- Modules transformed: 16.
- Build duration: 698 ms.
- JavaScript bundle: 927.96 kB.
- Gzip size: 266.49 kB.
- Non-blocking warning: bundle exceeds Vite's default 500 kB advisory threshold.

The public repository's own locked GitHub Actions build remains the authoritative integration gate after branch upload.

## Public testing boundary

Use only sanitized or synthetic GM and Stellantis workbooks in the public deployment. Do not upload production-derived warranty, claim, VIN, debit, supplier, dealer, or financial files to the public demo.

## Repository preservation

In the public demo repository:

1. Create a dedicated branch from current public `main`.
2. Rename the output artifact to `App.jsx`.
3. Replace only `src/App.jsx`.
4. Add this report if the public repository preserves implementation reports.
5. Run the existing public build workflow.
6. Review the diff before merging.
7. Open the deployed branch or merged public site and perform the browser acceptance checklist.
8. Keep the private authoritative repository unchanged.
