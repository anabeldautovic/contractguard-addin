# ContractGuard Word add-in (practice prototype)

A "Review with ContractGuard" button in Word. It compares the open supplier contract with a simplified OMV standard, checks each change against 8 draft rules, and adds a Word comment to each change. Fixed rules only, no AI model. The contract never leaves the computer.

- `taskpane.html` / `taskpane.js`: the side panel (uses Office.js to read the document and add comments)
- `rules.js`: the OMV standard text, the rulebook and the review logic
- `manifest.xml`: tells Word the add-in's name, button and web address
- `samples/`: supplier contracts to test with
- `INSTALL.md`: how to install it in Word

Hosted with GitHub Pages at https://anabeldautovic.github.io/contractguard-addin/

Practice data only: the OMV text paraphrases the public General Purchase Conditions (Nov 2025), and the rules are a draft, not approved by OMV Legal.
