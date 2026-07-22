# Photovoltaic System Recommender

A Python and Streamlit application developed as part of my bachelor's thesis at the Technical University of Cluj-Napoca.

## About the project

The application estimates a suitable residential photovoltaic system based on:

- household energy consumption;
- available budget;
- available roof area;
- photovoltaic panel characteristics;
- battery characteristics.

It compares the available equipment and generates indicative system recommendations according to the user's requirements.

## Main features

- photovoltaic panel and battery database;
- configuration based on user input;
- budget and roof-area constraints;
- comparison of suitable alternatives;
- export of generated recommendations.

## Technologies used

- Python
- Streamlit
- pandas
- openpyxl
- Microsoft Excel

## Project files

- `main.py` – main Streamlit application;
- `baze_date.xlsx` – photovoltaic panel and battery database;
- `requirements.txt` – required Python libraries.

## Running the application

Install the dependencies:
```bash
pip install -r requirements.txt
```

Run the application:

```bash
streamlit run main.py
```

## Author

**Flaviu Mariș**  
ETTI graduate – Technical University of Cluj-Napoca
```bash
pip install -r requirements.txt
