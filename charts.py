import os
import uuid
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")  # без вікон при збірці
import matplotlib.pyplot as plt
import seaborn as sns

def _tmp_png():
    return os.path.join(os.getcwd(), f"chart_{uuid.uuid4().hex[:8]}.png")

def select_and_build_charts(df, y, factors, results, selection: dict):
    imgs = []

    # гістограма/щільність за показником
    if selection.get("hist"):
        path = _tmp_png()
        plt.figure()
        s = pd.to_numeric(df[y], errors="coerce").dropna()
        sns.histplot(s, kde=True)
        plt.xlabel(y)
        plt.ylabel("Кількість")
        plt.title("Розподіл значень")
        plt.tight_layout()
        plt.savefig(path, dpi=130)
        plt.close()
        imgs.append(path)

    # box-plot за першим фактором
    if selection.get("box") and factors:
        path = _tmp_png()
        plt.figure()
        sns.boxplot(data=df, x=factors[0], y=y)
        plt.title(f"Box-plot: {y} за {factors[0]}")
        plt.tight_layout()
        plt.savefig(path, dpi=130)
        plt.close()
        imgs.append(path)

    # бар-чарт середніх ± SE
    if selection.get("bar") and "means" in results:
        means = results["means"].copy()
        path = _tmp_png()
        plt.figure()
        # якщо кілька факторів — візьмемо перший
        if factors:
            x = factors[0]
            grp = df.groupby(x)[y].agg(["mean", "sem"]).reset_index()
            plt.bar(grp[x].astype(str), grp["mean"])
            # помилки SE
            plt.errorbar(grp[x].astype(str), grp["mean"], yerr=grp["sem"], fmt="none")
            plt.xlabel(x)
        else:
            plt.bar(["Середнє"], [means["mean"].mean()])
        plt.ylabel(y)
        plt.title("Середні значення ± SE")
        plt.tight_layout()
        plt.savefig(path, dpi=130)
        plt.close()
        imgs.append(path)

    # лінія тренду (простий лінійний регрес)
    if selection.get("line"):
        # шукаємо якийсь числовий предиктор окрім y
        nums = [c for c in df.columns if c != y and pd.api.types.is_numeric_dtype(df[c])]
        if nums:
            x = nums[0]
            xvals = pd.to_numeric(df[x], errors="coerce")
            yvals = pd.to_numeric(df[y], errors="coerce")
            mask = xvals.notna() & yvals.notna()
            if mask.sum() >= 3:
                path = _tmp_png()
                plt.figure()
                plt.scatter(xvals[mask], yvals[mask])
                # тренд
                coef = np.polyfit(xvals[mask], yvals[mask], 1)
                poly = np.poly1d(coef)
                xs = np.linspace(xvals[mask].min(), xvals[mask].max(), 100)
                plt.plot(xs, poly(xs))
                plt.xlabel(x); plt.ylabel(y)
                plt.title(f"Лінія тренду: y={coef[0]:.3f}x+{coef[1]:.3f}")
                plt.tight_layout()
                plt.savefig(path, dpi=130)
                plt.close()
                imgs.append(path)

    # кругова — частки за першим фактором
    if selection.get("pie") and factors:
        path = _tmp_png()
        plt.figure()
        grp = df.groupby(factors[0])[y].count().reset_index()
        plt.pie(grp[y], labels=grp[factors[0]].astype(str), autopct="%1.1f%%")
        plt.title(f"Частки за {factors[0]}")
        plt.tight_layout()
        plt.savefig(path, dpi=130)
        plt.close()
        imgs.append(path)

    return imgs
