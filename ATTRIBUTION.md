# Attribution

nabla is a derivative work. The underlying function library and documentation workbook were created by **Craig Hatmaker** and published as the **Financial Starter Pack** for his BXL / 5g modelling ecosystem (Eloquens; the 6 July 2024 release, the latest located as at 18 August 2026). Copyright in the original work remains with the original author. This repository is not affiliated with or endorsed by him.

## Licence status

No open-source licence could be located for the upstream work as at 18 August 2026:

- The author's public gists carry a bare "Copyright" notice with no licence grant.
- The Beyond Excel site states no licence terms for the libraries.
- The Eloquens listing was not reachable for terms verification.

Absent a licence, the upstream author retains all rights in the original material. Treat the derived portions of this repository accordingly. If you are the upstream author and want changes to this repository, open an issue.

## What changed from the upstream workbook

- All namespaces renamed: `BXD`/`BXE`/`BXF`/`BXR`/`BXU`/`BXLDebt` were replaced by a single `nb.` prefix, with a one-letter tag where two modules shared a function name; 5g and BXL branding, branded artwork, the cover video thumbnails and their YouTube link, and Dropbox file links removed. One maths-citation link (a Diarmuid Early video, credited in an IntOnIntλ source comment) was retained deliberately.
- Help-block links that pointed at the upstream author's gists and site now point at this repository, relabelled from "Gist URL" to "Repository".
- The author's name is preserved in the revision histories inside the module sources (one upstream misspelling, "Hatmaekr", corrected), and the workbook's creator metadata credits him. Branding was removed; authorship records were not.
- American English converted to Australian English throughout, including function renames (`Amortizeλ` family to `Amortiseλ`).
- Calibri replaced with Aptos; US date formats replaced with day-first formats; sample data currency set to AUD.
- All sample and demonstration dates moved forward two years, calendar-aware (29 February maps to 28 February in non-leap targets); function version stamps set to 18 August 2026.
- Defects repaired: the undefined `Sheetλ` title formula on 46 worksheets, the undefined `BXE.Aboutλ` function, locale-fragile `RANDBETWEEN` text-date arguments, a dead table-of-contents link, and assorted typos.
- The foreign depreciation regime was removed outright, and five Australian functions were added: diminishing-value and prime-cost depreciation schedules, two GST helpers and a financial-year label. The two depreciation schedules are modelling helpers, not tax calculations, and v2.1.0 removed the claim that they implement an ATO method.
- Removed an empty Power Query data mashup, orphaned rich-value image residue, and the embedded printer configuration.
