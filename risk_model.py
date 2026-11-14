"""Python layer: logistic-regression attrition risk model."""
import numpy as np, pandas as pd
from sklearn.compose import ColumnTransformer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import roc_auc_score, accuracy_score, recall_score, precision_score
from sklearn.model_selection import train_test_split
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler

DROP = ["worker_id", "employee_number", "worker_status", "terminated", "age_band", "tenure_band", "comp_band"]

def fit_and_score(df: pd.DataFrame):
    X, y = df.drop(columns=[c for c in DROP if c in df]), df["terminated"]
    num = X.select_dtypes("number").columns.tolist()
    cat = X.select_dtypes(exclude="number").columns.tolist()
    pre = ColumnTransformer([("n", StandardScaler(), num), ("c", OneHotEncoder(handle_unknown="ignore"), cat)])
    pipe = Pipeline([("p", pre), ("m", LogisticRegression(max_iter=2000, class_weight="balanced", C=0.5))])
    Xtr, Xte, ytr, yte = train_test_split(X, y, test_size=0.25, random_state=42, stratify=y)
    pipe.fit(Xtr, ytr)
    pr, lab = pipe.predict_proba(Xte)[:, 1], pipe.predict(Xte)
    metrics = {"ROC AUC": roc_auc_score(yte, pr), "Accuracy": accuracy_score(yte, lab),
               "Recall (leavers)": recall_score(yte, lab), "Precision (leavers)": precision_score(yte, lab)}
    names = num + list(pipe.named_steps["p"].named_transformers_["c"].get_feature_names_out(cat))
    coef = pd.Series(pipe.named_steps["m"].coef_[0], index=names)
    coef = coef.reindex(coef.abs().sort_values(ascending=False).index).head(15).rename("coefficient").reset_index()
    coef.columns = ["feature", "coefficient"]
    pipe.fit(X, y)  # refit on all data for scoring
    scored = df.copy()
    scored["risk_score"] = (pipe.predict_proba(X)[:, 1] * 100).round(1)
    return metrics, coef, scored

def risk_flags(r) -> str:
    f = []
    if r.get("overtime") == "Yes": f.append("Overtime")
    if r.get("job_satisfaction", 4) <= 2: f.append("Low job satisfaction")
    if r.get("business_travel") == "Travel_Frequently": f.append("Frequent travel")
    if r.get("years_at_company", 9) <= 2: f.append("Short tenure")
    if r.get("years_since_promotion", 0) >= 5: f.append("Stalled promotion")
    if r.get("work_life_balance", 4) <= 2: f.append("Poor work-life balance")
    return ", ".join(f) or "-"
