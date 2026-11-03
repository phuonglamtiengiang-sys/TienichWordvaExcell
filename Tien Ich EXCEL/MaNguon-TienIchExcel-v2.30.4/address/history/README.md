# Historical address composition

Sources retrieved 2026-10-17:

- NSO public SOAP `https://danhmuchanhchinh.nso.gov.vn/DMDVHC.asmx`, `DanhMucPhuongXa`, snapshots 2024-01-01, 2025-01-01, 2025-06-30.
- NSO public comparison grid `https://danhmuchanhchinh.nso.gov.vn/Doi_Chieu_Moi.aspx`, changed rows between the above snapshots (1161 and 24). `fetch_nso.py` reproduces the documented UI's public read-only callback, reads every page and asserts the declared row count.
- `https://production.cas.so/address-kit/2024-01-01/communes`, for missing Bac Giang identities only. The response contains duplicated historical province assignments. The ten historical district codes select Bac Giang; old identities are retained only if they participate in a verified NQ1191 operation or a retained pre-July unit.
- NQ1191 full text: https://xaydungchinhsach.chinhphu.vn/toan-van-nghi-quyet-sap-xep-don-vi-hanh-chinh-cap-huyen-cap-xa-cua-tinh-bac-giang-119241008215240661.htm . The manual operation table corresponds to Articles 1/3 complete mergers and Article 2 Thanh Hai/Bien Son partial transfer. It yields 209 pre-change Bac Giang identities.
- Existing July mapping: tinhthanhpho.com 2026.10.8, CC BY 4.0.

`build_history.py` composes two dated stages and then joins the existing pre-July mapping by code, name, district and province. Comparisons supply explicit old/new identities and operative merger notes. A retained code alone does not establish footprint equivalence. Neighbor lists and office addresses are removed from source-name matching. Partial/remaining-area markers propagate review; paths without a link propagate unknown. Future July-decree notes attached to an earlier snapshot are not accepted as proof of a pre-July complete merger. Unknown branches remain even if another path has a known candidate.

Historical province names from the current SOAP join are sometimes incorrect; the legacy district/province catalog and historical comparison parents normalize them. Source paths and resolution references remain in `composition-audit.json`. `historical-additions.json` is the reproducible generated inventory; `unresolved-paths.json` records missing graph links. These are conservative machine-derived records, not a manual verification of every historical boundary.

Run `fetch_snapshots.py` and `fetch_nso.py` with the two date pairs to refresh raw sources, then run `build_sources.py` at the workspace root. Native Excel tests are in `tests/test_address.ps1`; they build v2.6 from v2.5.1, replace only the address module/data and version module, and remove temporary test helpers before saving.
