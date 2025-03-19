# Automation-Bot-for-french-geospatial-websites

This project is a Python-based automation bot designed to interact with five French geospatial websites. It automates the process of retrieving relevant PDFs based on user-provided input.

## Features

- Automates data retrieval from the following websites:
  - [Cadastre](https://cadastre.gouv.fr/scpc/accueil.do): Provides access to the French cadastral plan, comprising 598,133 plan sheets in image or vector formats.
  - [Géoportail de l'Urbanisme](https://www.geoportail-urbanisme.gouv.fr/map): Offers urban planning information and mapping services.
  - [PPRN Martinique](http://www.pprn972.fr/carto/web/): Provides access to natural risk prevention plans for Martinique.
  - [ERRIAL](https://errial.georisques.gouv.fr/#/): Allows users to quickly and easily assess the risks associated with their property.
  - [e-PLU Martinique](http://e-plu-martinique.com): An ArcGIS web application related to urban planning in Martinique.
- Utilizes Selenium for web automation.
- Features a Tkinter-based GUI for user input.
- Automatically downloads pdfs to your downloads directory.
