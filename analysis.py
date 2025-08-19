import pandas as pd
import numpy as np
from scipy.stats import shapiro, kruskal, mannwhitneyu, friedmanchisquare
import statsmodels.api as sm
import statsmodels.formula.api as smf
from statsmodels.stats.anova import anova_lm
import pingouin as pg
import scikit_posthocs as sp posthocs  # noqa  (для імпорту зберігаємо як posthocs)
# деякі середовища краще сприймають:  import scikit_posthocs as sp  і далі sp.posthoc_dunn
import scikit_posthocs as sp

def shapiro_check(df: pd.DataFrame, y: str):
    vals = pd.to_numeric(df[y], errors="coerce").dropna()
    if len(vals) < 3:
        return False, "Недостатньо даних для перевірки нормальності."
    stat, p = shapiro(vals)
    info = f"Шапіро–Вілка: W={stat:.3f}, p={p:.4f}."
    return (p >= 0.05), ("Дані сумісні з нормальним розподілом. " + info if p >= 0.05 else "Дані не нормальні. " + info)

# ---------- ПАРАМЕТРИЧНІ ----------
def _ols_anova(df, y, factors):
    """ANOVA через OLS формулу типу y ~ C(A) + C(B) + C(A):C(B) ..."""
    if not factors:
        return None, "Фактори не виявлено."
    terms = " + ".join([f"C({f})" for f in factors])
    # всі двовзаємодії
    inter = " + ".join([f"C({factors[i]}):C({factors[j]})" for i in range(len(factors)) for j in range(i+1, len(factors))])
    formula = f"{y} ~ {terms}" + (f" + {inter}" if inter else "")
    model = smf.ols(formula, data=df).fit()
    an = anova_lm(model, typ=2)
    return an, model

def _oneway_anova(df, y, factor):
    model = smf.ols(f"{y} ~ C({factor})", data=df).fit()
    an = anova_lm(model, typ=2)
    return an, model

def _twofactor_anova(df, y, f1, f2):
    model = smf.ols(f"{y} ~ C({f1}) + C({f2}) + C({f1}):C({f2})", data=df).fit()
    an = anova_lm(model, typ=2)
    return an, model

def _rm_anova(df, y, subject, within):
    # Pingouin RM ANOVA: вимагає стовпці: subject, within, dv
    d = df[[subject, within, y]].dropna()
    d = d.rename(columns={subject: "subject", within: "within", y: "dv"})
    res = pg.rm_anova(dv="dv", within="within", subject="subject", data=d, detailed=True)
    return res

def _lsd(df, y, factor):
    # LSD = парні t-тести без суворої корекції (класична НІР)
    # реалізація через pingouin.pairwise_ttests
    out = pg.pairwise_ttests(dv=y, between=factor, data=df, padjust="none")
    return out

def _tukey(df, y, factor):
    return pg.pairwise_tukey(dv=y, between=factor, data=df)

def _duncan(df, y, factor):
    # наближення: використовуємо scikit-posthocs Duncan
    return sp.posthoc_duncan(df, val_col=y, group_col=factor)

def _bonferroni(df, y, factor):
    return pg.pairwise_ttests(dv=y, between=factor, data=df, padjust="bonf")

def _effect_sizes(anova_table):
    # оцінка eta^2 (міра ефекту)
    tbl = anova_table.copy()
    if "sum_sq" in tbl.columns and "df" in tbl.columns:
        ss_total = tbl["sum_sq"].sum()
        tbl["eta2"] = tbl["sum_sq"] / ss_total if ss_total > 0 else np.nan
    return tbl

def run_parametric_suite(df: pd.DataFrame, y: str, factors, subject, sel: dict) -> dict:
    results = {}

    # базові усереднення
    if factors:
        grp = df.groupby(factors)[y].agg(["mean", "std", "count"]).reset_index()
        results["means"] = grp

    # однофакторний
    if sel.get("anova1") and factors:
        an, model = _oneway_anova(df, y, factors[0])
        results["anova1"] = an
        results["anova1_model"] = model

    # двофакторний
    if sel.get("anova2") and len(factors) >= 2:
        an, model = _twofactor_anova(df, y, factors[0], factors[1])
        results["anova2"] = an
        results["anova2_model"] = model

    # багато-факторний
    if sel.get("anovaN") and len(factors) >= 1:
        an, model = _ols_anova(df, y, factors)
        results["anovaN"] = an
        results["anovaN_model"] = model

    # RM-ANOVA
    if sel.get("rm") and subject and factors:
        res = _rm_anova(df, y, subject, factors[0])
        results["rm"] = res

    # Пост-хок по першому фактору (як правило так і роблять)
    if factors:
        f0 = factors[0]
        if sel.get("lsd"):
            results["lsd"] = _lsd(df, y, f0)
        if sel.get("tukey"):
            results["tukey"] = _tukey(df, y, f0)
        if sel.get("duncan"):
            results["duncan"] = _duncan(df, y, f0)
        if sel.get("bonf"):
            results["bonf"] = _bonferroni(df, y, f0)

    # сила впливу
    if sel.get("effect"):
        if "anovaN" in results:
            results["effect"] = _effect_sizes(results["anovaN"])
        elif "anova2" in results:
            results["effect"] = _effect_sizes(results["anova2"])
        elif "anova1" in results:
            results["effect"] = _effect_sizes(results["anova1"])

    return results


# ---------- НЕПАРАМЕТРИЧНІ ----------
def run_nonparametric_suite(df: pd.DataFrame, y: str, group: str, subject: str, sel: dict) -> dict:
    results = {}
    if group is None:
        results["info"] = "Груповий фактор для непараметричних тестів не виявлено."
        return results

    # Манна–Вітні (2 групи)
    if sel.get("mw"):
        levels = [g for g in df[group].dropna().unique()]
        if len(levels) == 2:
            a = pd.to_numeric(df[df[group] == levels[0]][y], errors="coerce").dropna()
            b = pd.to_numeric(df[df[group] == levels[1]][y], errors="coerce").dropna()
            stat, p = mannwhitneyu(a, b, alternative="two-sided")
            results["mannwhitney"] = pd.DataFrame({"stat": [stat], "p": [p], "groupA": [levels[0]], "groupB": [levels[1]]})
        else:
            results["mannwhitney"] = pd.DataFrame({"info": ["Потрібно рівно 2 групи."]})

    # Краскела–Уолліса
    if sel.get("kw"):
        groups = [pd.to_numeric(df[df[group] == lvl][y], errors="coerce").dropna() for lvl in df[group].dropna().unique()]
        if len(groups) >= 2:
            stat, p = kruskal(*groups)
            results["kruskal"] = pd.DataFrame({"stat": [stat], "p": [p], "k_groups": [len(groups)]})
        else:
            results["kruskal"] = pd.DataFrame({"info": ["Має бути ≥2 груп."]})

    # Фрідмана (RM)
    if sel.get("friedman") and subject:
        # очікуємо дані у форматі wide в межах суб'єкта; спростимо:
        # згрупуємо по subject та group, візьмемо середні, розгорнемо
        wide = df.pivot_table(index=subject, columns=group, values=y, aggfunc="mean")
        wide = wide.dropna()
        if wide.shape[1] >= 2:
            stat, p = friedmanchisquare(*[wide[c].values for c in wide.columns])
            results["friedman"] = pd.DataFrame({"stat": [stat], "p": [p], "k_levels": [wide.shape[1]]})
        else:
            results["friedman"] = pd.DataFrame({"info": ["Потрібно ≥2 рівнів фактора."]})

    # Пост-хок Данна (для KW)
    if sel.get("dunn"):
        try:
            ph = sp.posthoc_dunn(df, val_col=y, group_col=group, p_adjust="bonferroni")
            results["dunn"] = ph
        except Exception as e:
            results["dunn"] = pd.DataFrame({"error": [str(e)]})

    # базові узагальнення
    grp = df.groupby(group)[y].agg(["mean", "std", "count"]).reset_index()
    results["means"] = grp

    return results
