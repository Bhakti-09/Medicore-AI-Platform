import os
import pandas as pd

BASE_DIR = os.path.dirname(__file__)
DRUG_INTERACTIONS_PATH = os.path.join(BASE_DIR, 'drug_interaction_dataset_900.csv')

try:
    drug_interactions = pd.read_csv(DRUG_INTERACTIONS_PATH)
except FileNotFoundError:
    drug_interactions = pd.DataFrame(columns=['drug1','drug2','interaction','severity'])

def check_interaction(drug1, drug2):
    """Check if there's an interaction between drug1 and drug2."""
    if drug_interactions.empty or not drug1 or not drug2:
        return None, None
    mask = (
        ((drug_interactions['drug1'] == drug1) & (drug_interactions['drug2'] == drug2)) |
        ((drug_interactions['drug1'] == drug2) & (drug_interactions['drug2'] == drug1))
    )
    interaction = drug_interactions[mask]
    if not interaction.empty:
        return interaction.iloc[0]['interaction'], interaction.iloc[0]['severity']
    return None, None