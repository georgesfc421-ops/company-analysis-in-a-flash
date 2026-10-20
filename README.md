# Company Analysis in a Flash, The Future - Flash Analyst

Sales website for Flash Analyst, published with GitHub Pages. Everything in this repository is
public on the website.

- `*.html`: the pages
- `assets/site.css`, `assets/site.js`: shared styles and behaviour
- `assets/fonts/`: the Inter font, hosted here so no third party is contacted
- `assets/app/`: the browser app's engine (generated) and store settings (`config.js`)
- `assets/pyodide/`: Pyodide and openpyxl, which run the app in the visitor's browser
- `images/`, `samples/`: pictures, sample reports and the blank template

The pages, engine and security policy are generated from the main Flash Analyst repository
(`website/` folder, `tools/build_single_file.py`, `tools/add_security_meta.py`). Never put secrets,
customer data or license records in this repository.
