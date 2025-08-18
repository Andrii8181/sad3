import os
import pandas as pd
import matplotlib.pyplot as plt

def _safe_save(fig, out_dir, name):
    path = os.path.join(out_dir, name)
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)
    return path

def chart_means_by_factor(df, factor, dv, out_dir, name="means_by_factor.png"):
    g = df.groupby(factor)[dv].apply(pd.to_numeric, errors="coerce").agg(['mean','std','count'])
    fig = plt.figure()
    ax = fig.add_subplot(111)
    ax.bar(g.index.astype(str), g['mean'])
    ax.set_title(f"Середні значення за {factor}")
    ax.set_ylabel(dv)
    return _safe_save(fig, out_dir, name)

def chart_box_by_factor(df, factor, dv, out_dir, name="box_by_factor.png"):
    # простий boxplot без seaborn
    fig = plt.figure()
    ax = fig.add_subplot(111)
    groups = [pd.to_numeric(v, errors="coerce").dropna().values for k, v in df.groupby(factor)[dv]]
    labels = [str(k) for k in df.groupby(factor).groups.keys()]
    ax.boxplot(groups, labels=labels, showmeans=True)
    ax.set_title(f"BoxPlot для {dv} за {factor}")
    ax.set_ylabel(dv)
    return _safe_save(fig, out_dir, name)

def chart_histogram(df, dv, out_dir, name="histogram.png"):
    s = pd.to_numeric(df[dv], errors="coerce").dropna()
    fig = plt.figure()
    ax = fig.add_subplot(111)
    ax.hist(s, bins=10)
    ax.set_title(f"Гістограма {dv}")
    return _safe_save(fig, out_dir, name)

def build_default_charts_from_results(df, meta, results_dict, out_dir):
    paths = []
    dv = (meta["numeric"] or [df.columns[-1]])[0]
    # якщо є хоча б один фактор — зробимо бар і box
    if meta["factors"]:
        f = meta["factors"][0]
        paths.append(chart_means_by_factor(df, f, dv, out_dir, "01_means.png"))
        paths.append(chart_box_by_factor(df, f, dv, out_dir, "02_box.png"))
    # завжди — гістограма показника
    paths.append(chart_histogram(df, dv, out_dir, "03_hist.png"))
    return paths
