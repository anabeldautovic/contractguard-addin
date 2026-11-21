# ContractGuard Word add-in (practice prototype)

A "Review with ContractGuard" button in Word. It compares a supplier's returned contract with OMV's standard template, checks each change against a draft rulebook, and adds a Word comment to each change with a suggestion: Accept, Reject or Escalate. Changes no rule covers are marked "You decide". The lawyer decides every change with buttons in the side panel. Fixed rules only, no AI model. The contract never leaves the computer.

- `template.js`: the template library, word for word: OMV General Conditions of Purchase (Austrian law, January 2026 and November 2025). ContractGuard detects which one a supplier version is based on; a lawyer can also upload an earlier negotiation round to compare with. Public versions from [omv.com](https://www.omv.com/en/company/our-strength/supplier-center/general-conditions-of-purchase-at-omv)
- `rules.js`: the rulebook (9 draft rules) and the review logic
- `taskpane.html` / `taskpane.js` / `taskpane.css`: the side panel (uses Office.js to read the document, add comments and make tracked changes)
- `manifest.xml`: tells Word the add-in's name, button and web address
- `samples/`: `omv-template.docx` (the template), `supplier-example1.docx` (the task example: payment 60 → 30 days), `supplier-tough-round.docx` (11 realistic changes), `supplier-own-terms.docx` (not based on OMV’s template: ContractGuard stops), `supplier-nov2025-based.docx` (detected as Nov 2025), `round1-omv-sent.docx` + `round2-supplier.docx` (upload round 1 to compare round 2 with it)
- `INSTALL.md`: how to install it in Word

Hosted with GitHub Pages at https://anabeldautovic.github.io/contractguard-addin/

Practice only: the template is OMV's public text, and the rulebook is a draft, not approved by OMV Legal.
