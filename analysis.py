import numpy as np
import pandas as pd
from scipy.stats import shapiro, kruskal, friedmanchisquare, mannwhitneyu
import statsmodels.api as sm
from statsmodels.formula.api import ols
from statsmodels.stats.anova import AnovaRM
from statsmodels.stats.multicomp import pairwise_tukeyhsd
import scikit_posthocs as sp
import pingouin as pg

# --------- Службові ---------
def detect_columns(df: pd.DataFrame):
    # спробуємо здогадатися: числові vs фактори
    numeric = []
    factors = []
    for col in df.columns:
        s = pd.to_numeric(df[col], errors="coerce")
        # якщо >= 60% значень цифри — вважаємо числовим
        if s.notna().mean() >= 0.6:
            numeric.append(col)
        else:
            factors.append(col)
    # гарантовано хоча б 1 числова (візьмемо останню як показник)
    if not numeric and len(df.columns) > 0:
        numeric = [df.columns[-1]]
        factors = list(df.columns[:-1])
    return {"numeric": numeric, "factors": factors}

def check_normality(series: pd.Series):
    s = pd.to_numeric(series, errors="coerce").dropna()
    if len(s) < 3:
        # для Шапіро потрібно ≥3
        return 1.0
    stat, p = shapiro(s)
    return float(p)

def explain_non_normal(non_normal_cols, pvals):
    lines = [
        "Результат перевірки Шапіро–Вілка показав, що частина показників має ненормальний розподіл (p < 0.05).",
        "Колонки з відхиленням:"
    ]
    for c in non_normal_cols:
        lines.append(f" • {c}: p = {pvals[c]:.4f}")
    lines += [
        "",
        "Що робити науковцю?",
        " • Використати непараметричні методи (Kruskal–Wallis, Friedman, Dunn тощо);",
        " • або застосувати перетворення даних (логарифмування, квадратний корінь);",
        " • або збільшити обсяг вибірки/перевірити викиди."
    ]
    return "\n".join(lines)

def as_text_table(df: pd.DataFrame) -> str:
    with pd.option_context('display.max_rows', None, 'display.max_columns', None):
        return df.to_string()

# --------- Параметричні аналізи ---------
def anova_oneway(df, factors, numeric):
    # беремо перший фактор
    f = factors[0]
    y = numeric[0]
    model = ols(f"{y} ~ C({f})", data=df).fit()
    aov = sm.stats.anova_lm(model, typ=2)
    return {"ANOVA (1-фактор)": aov}

def anova_twoway(df, factors, numeric):
    f1, f2 = factors[:2]
    y = numeric[0]
    model = ols(f"{y} ~ C({f1}) + C({f2}) + C({f1}):C({f2})", data=df).fit()
    aov = sm.stats.anova_lm(model, typ=2)
    return {"ANOVA (2-фактори)": aov}

def anova_multi(df, factors, numeric):
    y = numeric[0]
    # всі взаємодії
    fterms = ":".join([f"C({f})" for f in factors])
    formula = f"{y} ~ " + " + ".join([f"C({f})" for f in factors]) + f" + {fterms}"
    model = ols(formula, data=df).fit()
    aov = sm.stats.anova_lm(model, typ=2)
    return {"ANOVA (багатофакторний)": aov}

def anova_rm(df, subject_col, within_factors, dv):
    # очікуємо, що в df уже є стовпець 'subject' (ідентифікатор повтору).
    aovrm = AnovaRM(df, depvar=dv, subject=subject_col, within=within_factors).fit()
    return {"ANOVA RM": aovrm.anova_table}

def lsd_posthoc(df, factor, dv):
    # Fisher LSD через pairwise t-tests без корекції (або з)
    res = pg.pairwise_ttests(dv=dv, between=factor, data=df, padjust='none')  # 'bonf' для Бонфероні
    return {"НІР (LSD)": res}

def tukey_posthoc(df, factor, dv):
    x = df[factor].astype(str)
    y = pd.to_numeric(df[dv], errors="coerce")
    ok = y.notna()
    tuk = pairwise_tukeyhsd(y[ok], x[ok])
    # у текст
    tuk_df = pd.DataFrame(data=tuk._results_table.data[1:], columns=tuk._results_table.data[0])
    return {"Тьюкі": tuk_df}

def duncan_posthoc(df, factor, dv):
    # через scikit_posthocs
    x = df[factor].astype(str)
    y = pd.to_numeric(df[dv], errors="coerce")
    data = pd.DataFrame({"group": x, "value": y})
    res = sp.posthoc_duncan(data, val_col="value", group_col="group")
    return {"Данкан": res}

def bonferroni_posthoc(df, factor, dv):
    res = pg.pairwise_ttests(dv=dv, between=factor, data=df, padjust='bonf')
    return {"Бонфероні": res}

def variance_report(df, numeric):
    y = pd.to_numeric(df[numeric[0]], errors="coerce").dropna()
    return {"Квадратичне відхилення": pd.DataFrame({"Середнє":[y.mean()], "Дисперсія":[y.var(ddof=1)], "Ст. відхилення":[y.std(ddof=1)]})}

def effect_size_eta(df, factor, dv):
    aov = pg.anova(data=df, dv=dv, between=factor, detailed=True)
    return {"Сила впливу (η²)": aov[["Source", "SS", "DF", "MS", "np2"]]}

# --------- Непараметричні ---------
def kruskal_oneway(df, factor, dv):
    groups = [pd.to_numeric(v, errors="coerce").dropna().values for k, v in df.groupby(factor)[dv]]
    stat, p = kruskal(*groups)
    return {"Kruskal–Wallis": pd.DataFrame({"H":[stat], "p":[p]})}

def friedman_test(df, subject, within, dv):
    # очікуємо довгий формат: на один subject — кілька рівнів within
    pivot = df.pivot(index=subject, columns=within, values=dv)
    stat, p = friedmanchisquare(*[pd.to_numeric(pivot[c], errors="coerce").dropna() for c in pivot.columns])
    return {"Friedman": pd.DataFrame({"Q":[stat], "p":[p]})}

def mann_whitney(df, factor, dv):
    # для двох груп
    levels = list(df[factor].astype(str).unique())
    if len(levels) != 2:
        return {"Mann–Whitney": pd.DataFrame({"p":["Потрібно рівно 2 групи"]})}
    a = pd.to_numeric(df[df[factor].astype(str)==levels[0]][dv], errors="coerce").dropna()
    b = pd.to_numeric(df[df[factor].astype(str)==levels[1]][dv], errors="coerce").dropna()
    stat, p = mannwhitneyu(a, b, alternative='two-sided')
    return {"Mann–Whitney": pd.DataFrame({"U":[stat], "p":[p]})}

def dunn_posthoc(df, factor, dv):
    x = df[factor].astype(str)
    y = pd.to_numeric(df[dv], errors="coerce")
    data = pd.DataFrame({"group": x, "value": y})
    res = sp.posthoc_dunn(data, val_col="value", group_col="group", p_adjust="bonferroni")
    return {"Dunn (пост-хок)": res}

# --------- Воркфлоу збірки результатів ---------
def run_selected_analyses(df, meta, methods):
    out = {}
    factors = meta["factors"] if meta["factors"] else [df.columns[0]]
    numeric = meta["numeric"] if meta["numeric"] else [df.columns[-1]]
    dv = numeric[0]

    # базові узгодження
    if len(factors) < 1 and any(k in methods for k in ["anova_oneway","anova_twoway","anova_multi","hsd_lsd","tukey","duncan","bonferroni","effect_size"]):
        out["Помітка"] = "Недостатньо факторів для факторного аналізу."

    if "anova_oneway" in methods and len(factors) >= 1:
        out |= anova_oneway(df, factors, numeric)

    if "anova_twoway" in methods and len(factors) >= 2:
        out |= anova_twoway(df, factors, numeric)

    if "anova_multi" in methods and len(factors) >= 2:
        out |= anova_multi(df, factors, numeric)

    if "anova_rm" in methods:
        # потребує колонки 'subject' та одного within-фактора (перший з factors)
        subj = "subject" if "subject" in df.columns else None
        within = factors[:1]
        if subj and within:
            out |= anova_rm(df, subject_col=subj, within_factors=within, dv=dv)
        else:
            out["ANOVA RM (примітка)"] = pd.DataFrame({"info":["Потрібна колонка 'subject' та within-фактор (перший фактор)"]})

    if "hsd_lsd" in methods and len(factors) >= 1:
        out |= lsd_posthoc(df, factors[0], dv)

    if "tukey" in methods and len(factors) >= 1:
        out |= tukey_posthoc(df, factors[0], dv)

    if "duncan" in methods and len(factors) >= 1:
        out |= duncan_posthoc(df, factors[0], dv)

    if "bonferroni" in methods and len(factors) >= 1:
        out |= bonferroni_posthoc(df, factors[0], dv)

    if "variance" in methods:
        out |= variance_report(df, numeric)

    if "effect_size" in methods and len(factors) >= 1:
        out |= effect_size_eta(df, factors[0], dv)

    return out

def run_selected_nonparametric_analyses(df, meta, methods):
    out = {}
    factors = meta["factors"] if meta["factors"] else [df.columns[0]]
    numeric = meta["numeric"] if meta["numeric"] else [df.columns[-1]]
    dv = numeric[0]

    if "kruskal" in methods and len(factors) >= 1:
        out |= kruskal_oneway(df, factors[0], dv)

    if "friedman" in methods:
        subj = "subject" if "subject" in df.columns else None
        within = factors[:1]
        if subj and within:
            out |= friedman_test(df, subject=subj, within=within[0], dv=dv)
        else:
            out["Friedman (примітка)"] = pd.DataFrame({"info":["Потрібна колонка 'subject' та один within-фактор"]})

    if "mannwhitney" in methods and len(factors) >= 1:
        out |= mann_whitney(df, factors[0], dv)

    if "dunn" in methods and len(factors) >= 1:
        out |= dunn_posthoc(df, factors[0], dv)

    if "variance" in methods:
        out |= variance_report(df, numeric)

    if "effect_size" in methods:
        # для непараметричних часто використовують ε² для KW — тут залишимо заголовок-заглушку
        out["Сила впливу (ε²)"] = pd.DataFrame({"примітка":["Оціночна метрика для непараметричних тестів; реалізуйте за потреби під ваш дизайн"]})

    return out
