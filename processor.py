import pandas as pd

def calculate_cvr(df):
    """
    df: DataFrame با ستون Item و ستون‌های R1 ... Rn
    خروجی: DataFrame با Item و CVR
    """
    result = []
    n = df.shape[1] - 1  # تعداد داورها
    
    for _, row in df.iterrows():
        item = row["Item"]
        ratings = row[1:]

        # فقط تعداد امتیازهای 3 (ضروری)
        ne = (ratings >= 3).sum()

        cvr = (ne - (n / 2)) / (n / 2)
        result.append([item, cvr])
    
    return pd.DataFrame(result, columns=["Item", "CVR"])


def calculate_cvi(df):
    """
    df: شامل Item، گروه Clarity_R*, Relevance_R*, Simplicity_R*
    خروجی: Item + CVI_Clarity + CVI_Relevance + CVI_Simplicity + CVI_Mean
    """

    def compute_single_cvi(cols):
        # تعداد امتیازهای 3 و 4
        df_scores = df[cols]
        return (df_scores >= 3).sum(axis=1) / df_scores.shape[1]

    # پیدا کردن ستون‌ها بر اساس نام‌ها
    clarity_cols = [c for c in df.columns if c.startswith("Clarity")]
    relevance_cols = [c for c in df.columns if c.startswith("Relevance")]
    simplicity_cols = [c for c in df.columns if c.startswith("Simplicity")]

    clarity = compute_single_cvi(clarity_cols)
    relevance = compute_single_cvi(relevance_cols)
    simplicity = compute_single_cvi(simplicity_cols)

    result = pd.DataFrame({
        "Item": df["Item"],
        "CVI_Clarity": clarity,
        "CVI_Relevance": relevance,
        "CVI_Simplicity": simplicity,
    })

    result["CVI_Mean"] = result[["CVI_Clarity","CVI_Relevance","CVI_Simplicity"]].mean(axis=1)

    return result
