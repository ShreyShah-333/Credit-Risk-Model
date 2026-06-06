
import pickle
import numpy as np
import pandas as pd

# Load model
MODEL_DIR = "/Users/shreyshah/Desktop/credit_risk_model/"

with open(MODEL_DIR + "xgb_model.pkl", "rb") as f:
    xgb_model = pickle.load(f)
with open(MODEL_DIR + "features.pkl", "rb") as f:
    FEATURES = pickle.load(f)
with open(MODEL_DIR + "calibration.pkl", "rb") as f:
    CALIBRATION = pickle.load(f)

def credit_score_to_rate(credit_score):
    if credit_score >= 800:   return 6.5
    elif credit_score >= 770: return 7.0
    elif credit_score >= 750: return 7.6
    elif credit_score >= 730: return 9.0
    elif credit_score >= 710: return 10.5
    elif credit_score >= 690: return 11.4
    elif credit_score >= 670: return 13.0
    elif credit_score >= 650: return 14.5
    elif credit_score >= 630: return 15.1
    elif credit_score >= 610: return 17.0
    elif credit_score >= 590: return 19.0
    elif credit_score >= 570: return 21.0
    elif credit_score >= 550: return 23.5
    elif credit_score >= 530: return 25.2
    else:                     return 27.0

def credit_score_to_grade_num(credit_score):
    if credit_score >= 750:   return 1
    elif credit_score >= 700: return 2
    elif credit_score >= 660: return 3
    elif credit_score >= 620: return 4
    elif credit_score >= 580: return 5
    elif credit_score >= 540: return 6
    else:                     return 7

def calibrate_pd(raw_pd):
    for low, high, cal_low, cal_high in CALIBRATION:
        if low <= raw_pd < high:
            ratio = (raw_pd - low) / (high - low)
            return cal_low + ratio * (cal_high - cal_low)
    return min(raw_pd, 0.65)

def predict_pd(borrower):
    annual_inc   = float(borrower.get("annual_income", 85000))
    monthly_debt = float(borrower.get("monthly_debt", 1800))
    loan_amnt    = float(borrower.get("loan_amount", 120000))
    credit_score = float(borrower.get("credit_score", 680))
    emp_years    = float(borrower.get("employment_years", 4.5))
    revol_util   = float(borrower.get("credit_util_pct", 0.45))
    inq_6m       = float(borrower.get("hard_inquiries", 2))
    delinq_2y    = float(borrower.get("late_payments", 1))
    pub_rec      = float(borrower.get("pub_rec", 0))
    open_acc     = float(borrower.get("open_acc", 8))
    total_acc    = float(borrower.get("total_acc", 20))
    revol_bal    = float(borrower.get("revol_bal", 10000))
    home         = borrower.get("home_ownership", "RENT")
    term         = int(borrower.get("loan_term_months", 60))
    purpose      = borrower.get("loan_purpose", "other")

    implied_rate = credit_score_to_rate(credit_score)
    grade_num    = credit_score_to_grade_num(credit_score)
    dti          = monthly_debt / max(annual_inc/12, 1) * 100
    installment  = loan_amnt / max(term, 1)
    dscr_proxy   = (annual_inc/12) / max(installment, 1)
    lti          = loan_amnt / max(annual_inc, 1)

    features = {
        "int_rate":          implied_rate,
        "grade_num":         grade_num,
        "term_months":       term,
        "dti":               dti,
        "emp_years":         emp_years,
        "is_mortgage":       1 if home in ["MORTGAGE","mortgage"] else 0,
        "is_rent":           1 if home in ["RENT","rent"] else 0,
        "inq_last_6mths":    inq_6m,
        "delinq_2yrs":       delinq_2y,
        "pub_rec":           pub_rec,
        "revol_util":        revol_util * 100 if revol_util <= 1 else revol_util,
        "open_acc":          open_acc,
        "total_acc":         total_acc,
        "dscr_proxy":        dscr_proxy,
        "loan_to_income":    lti,
        "log_annual_inc":    np.log1p(annual_inc),
        "log_revol_bal":     np.log1p(revol_bal),
        "log_loan_amnt":     np.log1p(loan_amnt),
        "high_risk_purpose": 1 if purpose in [
            "small_business","moving","medical",
            "renewable_energy","other"] else 0,
    }

    X        = pd.DataFrame([features])[FEATURES]
    raw_pd   = float(xgb_model.predict_proba(X)[0][1])
    cal_pd   = calibrate_pd(raw_pd)

    if cal_pd < 0.05:   grade = "AAA"
    elif cal_pd < 0.08: grade = "AA"
    elif cal_pd < 0.12: grade = "A"
    elif cal_pd < 0.18: grade = "BBB"
    elif cal_pd < 0.25: grade = "BB"
    elif cal_pd < 0.35: grade = "B"
    elif cal_pd < 0.45: grade = "CCC"
    else:               grade = "D"

    lgd            = 0.90
    credit_spread  = cal_pd * lgd * 1.5
    suggested_rate = round((0.0525 + credit_spread) * 100, 2)

    return {
        "pd":             round(cal_pd, 4),
        "pd_pct":         f"{cal_pd:.1%}",
        "grade":          grade,
        "suggested_rate": suggested_rate,
        "lgd":            lgd,
        "el":             round(cal_pd * lgd * loan_amnt, 0),
        "raw_pd":         round(raw_pd, 4),
    }
