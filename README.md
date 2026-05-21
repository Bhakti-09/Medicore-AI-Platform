# MediCore AI — Healthcare Advisor Platform

A full-stack AI-powered healthcare platform built with Streamlit, SQLite, and scikit-learn.

## Features

### Patient Portal
- 🔬 **AI Symptom Analysis** — RandomForest model predicts conditions from vitals + history
- 💊 **Drug Safety Check** — Real-time interaction detection against current medications
- 📅 **Appointment Booking** — Browse and book verified specialists
- 👨‍⚕️ **Doctor Directory** — Search by specialization, view profiles & fees
- 📁 **Medical History** — Full timeline with risk trends & charts
- 💊 **Prescription Tracker** — All AI and doctor-issued prescriptions in one place

### Doctor Portal
- 📅 **Appointment Management** — View, confirm, complete, and add notes to appointments
- 👥 **Patient Records** — See full medical history for patients who booked with you
- ✍️ **Write Prescriptions** — Issue prescriptions with auto drug-interaction warnings
- 📊 **Practice Analytics** — Appointment trends, completion rates
- 👤 **Profile Management** — Update availability, fee, hospital details

## Project Structure

```
.
├── app.py                          # Main Streamlit application
├── model.py                        # Train & save ML model
├── database.py                     # DB initialisation script
├── utils.py                        # Drug interaction helper
├── healthcare_main_dataset_900.csv # Primary dataset
├── drug_interaction_dataset_900.csv # Interaction dataset
├── disease_model.pkl               # Trained RF model (generated)
├── le_*.pkl                        # Label encoders (generated)
└── healthcare.db                   # SQLite database (generated)
```

## Setup

```bash
pip install streamlit pandas scikit-learn joblib

# 1. Train the model (run once)
python model.py

# 2. Initialise the database (optional, app.py does this automatically)
python database.py

# 3. Launch the app
streamlit run app.py
```

## Data Requirements

Ensure these CSV files are in the same directory:

**healthcare_main_dataset_900.csv** columns:
`age, gender, blood_pressure, blood_sugar, allergies, past_diseases, current_symptoms, previous_drugs, diagnosed_disease, recommended_drug, alternative_drug, side_effects, risk_level`

**drug_interaction_dataset_900.csv** columns:
`drug1, drug2, interaction, severity`

## Security Note
Passwords are SHA-256 hashed. For production, use bcrypt and move to PostgreSQL.