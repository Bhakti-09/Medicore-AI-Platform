import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import LabelEncoder
from sklearn.metrics import accuracy_score, classification_report, confusion_matrix
import joblib

# ─── Load Dataset ───────────────────────────────────────────────────────────────
df = pd.read_csv('healthcare_main_dataset_900.csv')

# ─── Encoders ───────────────────────────────────────────────────────────────────
le_gender           = LabelEncoder()
le_allergies        = LabelEncoder()
le_past_diseases    = LabelEncoder()
le_current_symptoms = LabelEncoder()
le_previous_drugs   = LabelEncoder()
le_diagnosed_disease= LabelEncoder()

df['gender_encoded']           = le_gender.fit_transform(df['gender'].astype(str))
df['allergies_encoded']        = le_allergies.fit_transform(df['allergies'].astype(str))
df['past_diseases_encoded']    = le_past_diseases.fit_transform(df['past_diseases'].astype(str))
df['current_symptoms_encoded'] = le_current_symptoms.fit_transform(df['current_symptoms'].astype(str))
df['previous_drugs_encoded']   = le_previous_drugs.fit_transform(df['previous_drugs'].astype(str))

# ─── Parse Blood Pressure ───────────────────────────────────────────────────────
df['systolic']  = df['blood_pressure'].apply(lambda x: int(str(x).split('/')[0]))
df['diastolic'] = df['blood_pressure'].apply(lambda x: int(str(x).split('/')[1]))

# ─── Features & Target ─────────────────────────────────────────────────────────
FEATURE_COLS = [
    'age', 'gender_encoded', 'systolic', 'diastolic', 'blood_sugar',
    'allergies_encoded', 'past_diseases_encoded',
    'current_symptoms_encoded', 'previous_drugs_encoded'
]

X = df[FEATURE_COLS]
y = le_diagnosed_disease.fit_transform(df['diagnosed_disease'])

# ─── Train / Test Split ─────────────────────────────────────────────────────────
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.2, random_state=42, stratify=y
)

# ─── Train ──────────────────────────────────────────────────────────────────────
model = RandomForestClassifier(n_estimators=200, max_depth=None, random_state=42, n_jobs=-1)
model.fit(X_train, y_train)

# ─── Evaluate ──────────────────────────────────────────────────────────────────
y_pred   = model.predict(X_test)
accuracy = accuracy_score(y_test, y_pred)
report   = classification_report(y_test, y_pred, target_names=le_diagnosed_disease.classes_)
confusion= confusion_matrix(y_test, y_pred)

print(f"\n✅ Model trained successfully.")
print(f"📊 Accuracy : {accuracy:.4f}")
print("\n📋 Classification Report:")
print(report)
print("🔢 Confusion Matrix:")
print(confusion)

# ─── Save Artifacts ────────────────────────────────────────────────────────────
joblib.dump(model,               'disease_model.pkl')
joblib.dump(le_gender,           'le_gender.pkl')
joblib.dump(le_allergies,        'le_allergies.pkl')
joblib.dump(le_past_diseases,    'le_past_diseases.pkl')
joblib.dump(le_current_symptoms, 'le_current_symptoms.pkl')
joblib.dump(le_previous_drugs,   'le_previous_drugs.pkl')
joblib.dump(le_diagnosed_disease,'le_diagnosed_disease.pkl')

print("\n💾 All model files saved.")